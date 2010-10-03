from distutils.core import setup
import py2exe
import glob
 
setup(console=['addin_assistant.pyw'],
      options={ "py2exe": { "dll_excludes": ["MSVCP90.dll"] }},
      data_files=[('images', glob.glob("images\\*.png")),
                  ('packaging', glob.glob("packaging\\*.*"))]
      )
