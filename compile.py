from support.compile import compile

compile("main_tk.py", "nuitka", console=False, icon="assets/sun_white.ico",
        plugin_qt=False, plugin_tk=True)

# !! WARN: add assets/sun_white.ico manually
