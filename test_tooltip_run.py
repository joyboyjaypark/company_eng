import time
import tkinter as tk
from threading import Thread

# Launch the app's main window from drawer.py
import importlib.util
import sys
spec = importlib.util.spec_from_file_location('drawer_mod', r'd:\company_eng\drawer.py')
drawer_mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(drawer_mod)

# create the app root and instance
root = tk.Tk()
app = drawer_mod.ResizableRectApp(root)

# helper to create a minimal room polygon and force auto-place
def setup_and_test():
    rc = app.get_current_palette()
    # create a simple rect shape (x1,y1,x2,y2)
    s = rc.create_rect_shape(50, 50, 250, 150, editable=True, color='black')
    # run auto-generate labels and flows
    try:
        rc.generated_space_labels.clear()
        rc.auto_generate_current()  # if exists; otherwise call compute_and_apply_supply_flow then auto_place_diffusers
    except Exception:
        # fallback: compute flow then place diffusers
        try:
            rc.compute_and_apply_supply_flow()
        except Exception:
            pass
        try:
            rc.auto_place_diffusers(10.0)
        except Exception:
            pass

    # find a diffuser id
    did = None
    for lab in rc.generated_space_labels:
        if lab.get('diffuser_ids'):
            did = lab['diffuser_ids'][0]
            break

    if did is None:
        print('No diffuser created; exiting')
        root.quit()
        return

    # get center
    coords = rc.canvas.coords(did)
    cx = (coords[0]+coords[2])/2
    cy = (coords[1]+coords[3])/2

    # simulate mouse motion events near the center
    for dx in range(-5, 6, 5):
        try:
            event = tk.Event()
            event.x = int(cx + dx)
            event.y = int(cy)
            rc.on_mouse_move(event)
            # allow UI update
            root.update_idletasks()
            root.update()
            time.sleep(0.2)
            if getattr(rc, 'flow_tooltip_id', None):
                print('Tooltip id:', rc.flow_tooltip_id)
                break
        except Exception as e:
            print('Exception during motion simulation:', e)
            break

    # keep the window for a short time so manual inspection possible
    time.sleep(1.0)
    root.quit()

# run setup in a separate thread to avoid blocking tkinter mainloop
Thread(target=setup_and_test).start()
root.mainloop()
print('Script finished')
