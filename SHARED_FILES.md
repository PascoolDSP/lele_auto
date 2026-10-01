# Shared files for project lele_auto

Pascal can drop documents for this project from the file-drop site
**https://partage.home/** (or the raw LAN URL http://192.168.1.100:8129/ for
files larger than 50 MB, which bypasses the nginx cap). He selects **lele_auto**
there and the files land locally in this project at:

    ./share/

## For Claude (this session)

- When Pascal says he "shared", "dropped" or "sent" a file/document for this
  project, look in `./share/` — that is where they arrive.
- Treat `./share/` as the project inbox: read, reference or process those files
  from there. It sits outside `site/`, so it is NOT served on the web.
- Uploads keep their original name; duplicates get a ` (1)`, ` (2)` … suffix.
- The folder is local to the project. Safe to read; ask before deleting.
