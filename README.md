# Introduction
Watcher for ActivityWatch which:
- when Outlook is the active window (specifically a process called outlook.exe)
- gets the name of the selected item, if any
- logs the name to ActivityWatch, along with the folder (inbox, sent, Calendar)

This could be the email Subject, Calendar event title (both tested - other items might also show up).

Only works with Windows platform.
Only tested with the "old" Outlook client.

# How to test
1. Create and activate a VirtualEnv
2. Clone the repo to the venv
3. Install dependencies: `pip install aw-core aw-client pywin32`
4. Optionally: `pip install pyinstaller`
5. Start an AW server - in testing mode if you wish
    1. If you have a regular installation of AW (not dev), you can do this:
    2. Search for the aw-server.exe executable. In my case this was in `C:\Users\USERNAME\AppData\Local\Programs\ActivityWatch\aw-server\`
    3. Run with `aw.server.exe --testing`
6. Run: `python -m aw_watcher_outlook --testing`
7. By default it print the item name (or `none` if outlook is not active, or if no item is selected). Browse through outlook to see the results.
8. Check the timeline by browsing to `localhost:5666`

# Repo structure
1. This repo was mostly vibe-coded, with additional copying from AFK and Window watchers. The results is a third way of setting up a watcher.
2. These instructions are for noobs, like myself.
3. `__init__.py` makes this a python module.
4. `__main__.py` is the entry point when running the module.
5. `outlook.py` is the core AW client / loop.
6. `windows.py` is the engine which gets item names - using COM via pywin32.
7. `config.py` is copied and amended from AFK.

# How it works
1. `__main__.py` is the entry point and calls the main function in outlook.py
2. The main function:
    1. Loads the arguments
    2. Initializing logging
    3. Initializes the client
    4. Creates the bucket if needed
    5. Starts the main loop
3. The main loop:
    1. Checks if the active process is `outlook.exe`
    2. Retrieves the item name
    3. Sends the event/heartbeat to the AW server

# Integrating to your setup
This is useful especially if you have a "traditionally installed" AW setup.

1. Test as per the above.
2. When tests are successful, run pyinstaller: `pyinstaller .\aw-watcher-outlook.spec --clean --noconfirm`
3. In the newly created `dist` directory, there is another directory called `aw-watcher-outlook`
4. Copy this to the same location as the other watchers. In my case it is `C:\Users\USERNAME\AppData\Local\Programs\ActivityWatch\`
5. (Unless AW does some cool stuff itself,) create the config file in `C:\Users\USERNAME\AppData\Local\activitywatch\activitywatch\aw-watcher-outlook\`
    ```
    # aw-watcher-outlook.toml
    [aw-watcher-outlook]
    poll_time = 2

    [aw-watcher-outlook-testing]
    poll_time = 1
    ```
6. Restart ActivityWatch - aw-watcher-outlook should be available in the list of modules from the system tray icon.
