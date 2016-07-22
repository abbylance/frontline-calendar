from distutils.core import setup
import py2exe

setup(name='frontline-calendar',
      scripts=['frontline_calendar.py'],
      console=['frontline_calendar.py'],
      options={'py2exe': {'excludes': ['six.moves.urllib.parse']}},
      data_files=[('certs', ['certs/cacerts.txt']),
                  ('zoneinfo', ['zoneinfo/Chicago'])])
