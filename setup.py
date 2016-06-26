from distutils.core import setup
import py2exe

setup(name='frontline-calendar',
      scripts=['frontline-calendar.py'],
      console=['frontline-calendar.py'],
      options={'py2exe': {'excludes': ['six.moves.urllib.parse']}},
      data_files=[('certs', ['certs/cacerts.txt']),
                  ('zoneinfo', ['zoneinfo/Chicago'])])
