# excel-link-restorer
Restores broken Excel links, Explanation TBD

# .bas File -> Excel macro

Run the first sub in your Excel using the IDE found via **Alt + F11**. "Localizes" all Excel links, and therefore removes all user-specific information or broken links that are mixtures of Web-file:/// URIs and local drives (D:\), e.g. file:///D:\

# .sh File -> Bash script
Much better solution and much faster and leaner. Runs in bash on Windows or Linux natively. Be sure to install the package fuse-zip from your repostitory. Also needs perl to be installed.
beforehand, run
`sudo ln -s /proc/self/mounts /etc/mtab`

Just start it with `./doit.sh`
