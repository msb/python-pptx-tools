"""
This script demonstrates how you can use the `python-pptx` library
(https://github.com/scanny/python-pptx) to edit the animation in a Power Point file.
Out of the box, python-pptx supports the basic presentation elements but does not support
animation. This script illustrates how you can use the library to edit the xml directly thus giving
you full control over any part of the file.

This particular example double the speed of a particular animation in the 1st slide.

UPDATE: This is actually a lot easier than I thought.
`pptx` provides a `pptx.slides[x].element` attribute
that proxies the underlying xml for a slide. You can use this to edit the xml directly.
However, this approach might come in handy so leaving it as example code.

Usage:
  editor-poc.py <path_to_file>
  editor-poc.py (-h | --help)
  editor-poc.py --version

Options:
  -h --help  this screen.
  --version  Show version.
"""

from pathlib import Path

from docopt import docopt
from pptx import Presentation
from pptx.oxml.ns import nsuri

# the "p" xml namespace
p = "{" + nsuri('p') + "}"

def edit(path_to_file):

    prs = Presentation(path_to_file)

    for part in prs.part.package.iter_parts():
        if part._partname == "/ppt/slides/slide1.xml":
            timing_el = next(child for child in part._element if child.tag == f"{p}timing")
            for cond in timing_el.findall(f".//{p}cond"):
                try:
                    delay = int(cond.get("delay"))
                    if delay > 0:
                        cond.set("delay", str(delay / 2))
                except ValueError:
                    pass
    
    prs.save(Path(path_to_file).with_suffix(".edited.pptx"))

if __name__ == "__main__":
    args = docopt(__doc__, version="0.1")
    edit(args["<path_to_file>"])
