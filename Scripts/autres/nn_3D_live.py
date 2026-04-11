import numpy as np
import torch
import torch.nn as nn
import pyvista as pv

# -----------------------
# CONFIG
# -----------------------
SEED = 0
np.random.seed(SEED)
torch.manual_seed(SEED)

layer_sizes = [16, 64, 64, 64, 32, 16, 8]
batch_size = 512
lr = 1e-3

edge_sample_per_node = 4
max_edges_total = 200000      # si ça rame: 6000
update_ms = 35
sphere_radius = 0.16

# -----------------------
# DATASET jouet
# -----------------------
N = 8000
X = torch.randn(N, layer_sizes[0])
Wtrue = torch.randn(layer_sizes[0], 1)
Y = torch.sin(X @ Wtrue) + 0.1 * torch.randn(N, 1)

# -----------------------
# MODELE
# -----------------------
class BigMLP(nn.Module):
    def __init__(self, sizes):
        super().__init__()
        layers = []
        for i in range(len(sizes)-1):
            layers.append(nn.Linear(sizes[i], sizes[i+1]))
            if i < len(sizes)-2:
                layers.append(nn.ReLU())
        self.net = nn.Sequential(*layers)

    def forward(self, x):
        return self.net(x)

model = BigMLP(layer_sizes)
opt = torch.optim.Adam(model.parameters(), lr=lr)
loss_fn = nn.MSELoss()

linear_layers = [m for m in model.modules() if isinstance(m, nn.Linear)]
offsets = np.cumsum([0] + layer_sizes[:-1])
total_nodes = int(sum(layer_sizes))

# hooks activations
acts = [None] * len(linear_layers)
def make_hook(i):
    def _h(m, inp, out):
        acts[i] = out.detach()
    return _h
hooks = [lin.register_forward_hook(make_hook(i)) for i, lin in enumerate(linear_layers)]

# -----------------------
# POSITIONS 3D (spirale ciné)
# -----------------------
points = np.zeros((total_nodes, 3), dtype=np.float32)
gid = 0
turns = 9.0
for li, sz in enumerate(layer_sizes):
    t = np.linspace(0, 1, sz, endpoint=False)
    ang = 2*np.pi*(t*turns + li*0.15)
    r = 2.0 + 0.25*li
    x0 = li * 2.3
    for j in range(sz):
        arm = -1.0 if (j % 2 == 0) else 1.0
        x = x0 + 0.10*np.cos(ang[j]*0.9)
        y = arm * r*np.cos(ang[j]) * 0.55
        z = r*np.sin(ang[j]) * 0.55
        points[gid] = (x, y, z)
        gid += 1

# -----------------------
# EDGES échantillonnés
# -----------------------
edge_src, edge_dst = [], []
for li, lin in enumerate(linear_layers):
    W = lin.weight.detach().cpu().numpy()
    in_sz = W.shape[1]
    out_sz = W.shape[0]
    src_base = offsets[li]
    dst_base = offsets[li+1]
    for j in range(in_sz):
        col = np.abs(W[:, j])
        k = min(edge_sample_per_node, out_sz)
        topk = np.argpartition(-col, kth=k-1)[:k]
        for oi in topk:
            edge_src.append(src_base + j)
            edge_dst.append(dst_base + oi)

edge_src = np.array(edge_src, dtype=np.int32)
edge_dst = np.array(edge_dst, dtype=np.int32)

if len(edge_src) > max_edges_total:
    idx = np.random.choice(len(edge_src), size=max_edges_total, replace=False)
    edge_src = edge_src[idx]
    edge_dst = edge_dst[idx]

E = len(edge_src)

# mapping edges -> (layer, inj, outi)
edge_map_layer = np.zeros(E, dtype=np.int32)
edge_map_inj = np.zeros(E, dtype=np.int32)
edge_map_outi = np.zeros(E, dtype=np.int32)

for li, lin in enumerate(linear_layers):
    in_sz = lin.in_features
    out_sz = lin.out_features
    src_base = offsets[li]
    dst_base = offsets[li+1]
    mask = (edge_src >= src_base) & (edge_src < src_base + in_sz)
    mask &= (edge_dst >= dst_base) & (edge_dst < dst_base + out_sz)
    idxs = np.where(mask)[0]
    if idxs.size:
        edge_map_layer[idxs] = li
        edge_map_inj[idxs] = edge_src[idxs] - src_base
        edge_map_outi[idxs] = edge_dst[idxs] - dst_base

def edge_strength():
    s = np.zeros(E, dtype=np.float32)
    for li, lin in enumerate(linear_layers):
        W = lin.weight.detach().abs().cpu().numpy()
        idxs = np.where(edge_map_layer == li)[0]
        if idxs.size == 0:
            continue
        inj = edge_map_inj[idxs]
        outi = edge_map_outi[idxs]
        s[idxs] = W[outi, inj]
    s /= (s.max() + 1e-9)
    return s

def node_activations(xb):
    c = np.zeros(total_nodes, dtype=np.float32)

    v0 = xb.detach().abs().mean(dim=0).cpu().numpy()
    v0 = v0 / (v0.max() + 1e-9)
    c[offsets[0]:offsets[0]+layer_sizes[0]] = v0

    for li, a in enumerate(acts):
        if a is None:
            continue
        v = a.detach().abs().mean(dim=0).cpu().numpy()
        v = v / (v.max() + 1e-9)
        c[offsets[li+1]:offsets[li+1]+layer_sizes[li+1]] = v

    return np.clip(c**0.55, 0, 1)

# -----------------------
# EDGES mesh robuste: poids en point_data
# -----------------------
def build_edges_pointdata(src, dst, w):
    e = len(src)
    if e == 0:
        m = pv.PolyData(np.zeros((0, 3), dtype=np.float32))
        m.lines = np.array([], dtype=np.int64)
        m.point_data["w"] = np.array([], dtype=np.float32)
        return m

    pts = np.empty((2*e, 3), dtype=np.float32)
    pts[0::2] = points[src]
    pts[1::2] = points[dst]

    lines = np.empty((3*e,), dtype=np.int64)
    lines[0::3] = 2
    lines[1::3] = np.arange(0, 2*e, 2, dtype=np.int64)
    lines[2::3] = np.arange(1, 2*e, 2, dtype=np.int64)

    m = pv.PolyData(pts)
    m.lines = lines
    m.point_data["w"] = np.repeat(w.astype(np.float32), 2)
    return m

def edge_tiers(w):
    q1, q2 = np.quantile(w, [0.70, 0.92])
    m1 = w <= q1
    m2 = (w > q1) & (w <= q2)
    m3 = w > q2
    return m1, m2, m3

# -----------------------
# BUILD SCENE
# -----------------------
pv.set_plot_theme("dark")
plotter = pv.Plotter(window_size=(1400, 850))
plotter.set_background("black")

# nodes spheres
node_cloud = pv.PolyData(points)
node_cloud["act"] = np.zeros(total_nodes, dtype=np.float32)
sphere = pv.Sphere(radius=sphere_radius, theta_resolution=16, phi_resolution=16)
node_glyphs = node_cloud.glyph(scale=False, geom=sphere, orient=False)
node_glyphs["act"] = np.zeros(node_glyphs.n_points, dtype=np.float32)

# initial edges
w0 = edge_strength()
m1, m2, m3 = edge_tiers(w0)
e1 = build_edges_pointdata(edge_src[m1], edge_dst[m1], w0[m1])
e2 = build_edges_pointdata(edge_src[m2], edge_dst[m2], w0[m2])
e3 = build_edges_pointdata(edge_src[m3], edge_dst[m3], w0[m3])

a1 = plotter.add_mesh(e1, scalars="w", cmap="turbo", opacity=0.08, line_width=1, lighting=False)
a2 = plotter.add_mesh(e2, scalars="w", cmap="turbo", opacity=0.18, line_width=2, lighting=False)
a3 = plotter.add_mesh(e3, scalars="w", cmap="turbo", opacity=0.35, line_width=4, lighting=False)

plotter.add_mesh(node_glyphs, scalars="act", cmap="turbo", opacity=0.98, specular=0.7, smooth_shading=True)

plotter.add_text("NN 3D Live (PyVista) — Training en cours", color="white", font_size=12, position="upper_left")
plotter.add_text("step 0 — loss ...", color="white", font_size=12, position="lower_left", name="hud")

plotter.camera_position = "iso"
plotter.camera.zoom(1.4)

state = {"step": 0, "rebuild_every": 10, "loss": 0.0}

def rebuild_edges(w):
    m1, m2, m3 = edge_tiers(w)
    new1 = build_edges_pointdata(edge_src[m1], edge_dst[m1], w[m1])
    new2 = build_edges_pointdata(edge_src[m2], edge_dst[m2], w[m2])
    new3 = build_edges_pointdata(edge_src[m3], edge_dst[m3], w[m3])
    a1.mapper.dataset.shallow_copy(new1)
    a2.mapper.dataset.shallow_copy(new2)
    a3.mapper.dataset.shallow_copy(new3)

def update():
    # train
    idx = torch.randint(0, N, (batch_size,))
    xb = X[idx]
    yb = Y[idx]

    model.train()
    opt.zero_grad()
    pred = model(xb)
    loss = loss_fn(pred, yb)
    loss.backward()
    opt.step()
    state["loss"] = float(loss.item())

    # eval for hooks
    model.eval()
    with torch.no_grad():
        _ = model(xb)

    # nodes colors
    c = node_activations(xb)
    node_glyphs["act"] = np.interp(
        np.linspace(0, total_nodes-1, node_glyphs.n_points),
        np.arange(total_nodes),
        c
    ).astype(np.float32)

    # edges
    w = edge_strength()
    if state["step"] % state["rebuild_every"] == 0:
        rebuild_edges(w)
    else:
        avg = float(w.mean())
        a1.prop.opacity = 0.05 + 0.10 * avg
        a2.prop.opacity = 0.12 + 0.18 * avg
        a3.prop.opacity = 0.22 + 0.30 * avg

    # camera rotation + hud
    plotter.camera.azimuth += 0.7
    plotter.add_text(f"step {state['step']} — loss {state['loss']:.4f}",
                     color="white", font_size=12, position="lower_left", name="hud")
    state["step"] += 1

# -----------------------
# RUN with VTK timer (safe: prefer add_callback, fallback to iren with guards)
# -----------------------
try:
    # First try the high-level API: add_callback is safer and registers before show()
    used_cb = False
    try:
        # plotter.add_callback returns an id in recent pyvista versions
        cb_id = plotter.add_callback(update, interval=update_ms)
        print(f"[PYVISTA] registered add_callback id={cb_id}")
        used_cb = True
    except Exception:
        # older pyvista / backend might not support add_callback: we'll fallback below
        used_cb = False

    # Show the window. If the user closes the window immediately, plotter.iren may be None.
    plotter.show(auto_close=False)

    if not used_cb:
        # Fallback: try to register a repeating timer on the interactor if available
        iren = getattr(plotter, "iren", None)
        if iren is None:
            # Window was closed (or interactor not created). Skip timed updates gracefully.
            print("[PYVISTA] interactor not available after show(); skipping timer registration.")
        else:
            def _timer_cb(obj, event):
                update()

            try:
                iren.add_observer("TimerEvent", _timer_cb)
                iren.create_repeating_timer(update_ms)
                iren.start()
            except Exception as e:
                print("[PYVISTA] failed to register interactor timer:", e)
                # nothing else to do; allow program to continue/exit

finally:
    # remove hooks cleanly; ignore errors during teardown
    for h in hooks:
        try:
            h.remove()
        except Exception:
            pass
