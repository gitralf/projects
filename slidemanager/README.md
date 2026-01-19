# Slidemanager

  - Do you have a lot of PPTX from where you want to reuse some slides?
  - Do you struggle to find the right presentation when looking for a specific slide?
  - Do you sometimes want to combine single slides from different presentations into a new one?

  These scripts might be able to help you...

## Make-Thumbs.ps1
  Most important is Make-Thumb: It will create for each presentation in a directory a PNG for each slide and name this with a combination of the original pptx filename and slide number. All PNG will be stored in a flat directory to make it easy to scroll through. Simply

- give the Path where your PPTX files are stored
- give the OutPath where Thumbs should be stored

and run. Example:

<code>.\Make-Thumb.ps1 -Path c:\temp\allPPTX -OutPath c:\temp\allThumbs</code>

The script will open every PowerPoint presentation in <code>c:\temp\allPPTX</code> and will export for each slide a PNG file into <code>c:\temp\allThumbs</code>. It will only process the PPTX when the last modified date is newer than the slide-PNGs, otherwise will skip the presentation.

This script is the base for other slidemanager-scripts. It will produce a logfile in the OutPath.

## Find-Duplicate.ps1

The script will build a MD5 hash for each PNG in the given directory ans will list duplicates. If you enter a <code>-MovetoPath</code> it will also move the duplicates to the given folder.

## Make-Newpptx.ps1

This script will read all PNGs from a given directory and present them in a scalable window. Each image has a checkbox, and you can simply select all the images that you want. When finished, click the "Export" button on the top of the window, and the script will create a new presentation by combining the original slides from the different PPTX files into a new PPTX file.

Usage is simple:

<code>Make-Newpptx.ps1 -ThumbsPath c:\temp\allThumbs -PPTXPath c:\temp\allPPTX</code>

You need to provide both the directory with the original PPTX and the directory with the thumbs generate with <code>Make-Thumb.ps1</code>.



