"""
This script extract images from a powerpoint presentation.
If `<search_string>` is provided, only images are extracted whose name matches the search string.
The script will also do a dry run.

Usage:
  extract-images.py <path_to_file> [<search_string>] [(-d | --dryrun)]
  extract-images.py (-h | --help)
  extract-images.py --version

Options:
  <path_to_file>  path to the Power Point file.
  <search_string> name of image to search for in the presentation (regex).
  -d --dryrun     Show what would be extracted without actually extracting.
  -h --help       this screen.
  --version       Show version.
"""

import re

from docopt import docopt
from pptx import Presentation
from slugify import slugify

def extract_images(path_to_file, search_string, dryrun=False):

  def do(container_name, shapes):
    """
    Recursively extract images from the presentation shapes.
    If a search string is provided, only extract images whose name matches the search string.
    """
    for shape in shapes:
      name = container_name + "-" + shape.name
      if hasattr(shape, "shapes"): # `shape` is a grouping of sub-shapes
        do(name, shape.shapes)
      elif hasattr(shape, "image"):
        searchable = slugify(name)
        if search_string is None or re.search(search_string, searchable):
          filename = searchable + "." + shape.image.ext
          if dryrun:
            print(f"Would extract '{filename}'")
          else:
            with open(filename, "wb") as image_file:
              image_file.write(shape.image.blob)
            print(f"Extracted '{filename}'")

  prs = Presentation(path_to_file)

  for i, slide in enumerate(prs.slides):
    do(f"slide {i}", slide.shapes)


if __name__ == "__main__":
  args = docopt(__doc__, version="0.1")
  extract_images(args["<path_to_file>"], args.get("<search_string>"), dryrun=args.get("--dryrun"))
