# latex-in-outlook
This repo contains all you need to start typing LaTeX equations in Office Outlook (OWA).

The following instructions are for UNIX users. Alternatively you can do it on Windows with WSL and systemd enabled in your wsl.conf file.

First, clone the repo onto your local machine. Ensure you have python3 and the python library Flask installed. 
You will also want to make sure you have a installation of pdflatex and ghostscript. You need to change the absolute paths to these executables in the python file.
You then can create a systemd service to ensure that the flask server always runs by placing the .service file in a suitable place and execute the commands:

```
$ systemctl --user daemon-reload
$ systemctl --user enable --now latex-render.service
$ systemctl --user status latex-render.service
```
You should see that the service is running (active). You can also check if it works by:
```
$ curl http://127.0.0.1:8765/health
```
If something has gone amiss, you can debug with:
```
$ journalctl --user -u latex-render.service -b -n 200 --no-pager
```

Now, you can install <https://www.tampermonkey.net/> Tampermonkey on your browser.

You can then import the .js script into your Tamper monkey. Upon enabling and refreshing the Outlook page, you should see some pill buttons in the editing area of the text. Enclose your maths in single/double dollar signs and click the PNG buttons, this will replace the maths equations with images of the rendered TeX equations.

