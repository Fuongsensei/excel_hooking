
import threading as th

foreground : th.Event = th.Event()
is_done :th.Event = th.Event()
wait_printing = th.Event()
is_printing  = th.Event()
is_alt_l = th.Event()
