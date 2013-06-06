from distutils.core import setup
import py2exe
import glob
 
setup(windows=[
                {'script': 'addin_assistant.pyw',
                 'icon_resources': [(1, "images\\AddInDesktop.ico")]
                }
              ],
      options={ "py2exe": { "dll_excludes": ["MSVCP90.dll"] }},
      data_files=[('images', glob.glob("images\\*.png") + 
                             glob.glob("images\\*.ico")),
                  ('packaging', glob.glob("packaging\\*.*")),
                  ('resources', glob.glob("resources\\*.*"))]
      )
