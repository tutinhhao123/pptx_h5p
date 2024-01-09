Powerpoint to .h5p course presentation converter. Windows (+ MS Office) only.

-Instructions
Download the executable of this tool
As user @dgcruzing described it in issue #1:
Create a shortcut on your desktop to it. i.e right click on "ppt2h5p" -> sendto -> Desktop
Open your explorer, type shell:sendto into the addressbar
hit enter, this should take you to your users/yourusername/Appdata/roaming/Microsoft/windows/Sendto folder
drag the shortcut here you have created
Right-click on your .pptx file, send to "ppt2h5p". Let the magic happen.
Upload .h5p into your H5P repository

-How it works
pptx2h5p makes use of the COM interface offered by an installed powerpoint on the host.

opens the .pptx file in powerpoint
exports each slide as .png file
analyses images for width/height to adjust scaling for h5p
imports the .png files into an .hp5 archive (it's a zip essentially)

