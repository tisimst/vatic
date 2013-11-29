from vatic import *

v = vatic()
v.setnpts(1000)
# v.verbose = True

v.addassumption(cell='C8', tag=v.getcell(cell='B8'), 
    dist={N: (v.getcell(cell='C8'), v.getcell(cell='D8')/3)})

v.addforecast(cell='C30', tag=v.getcell(cell='B30'), 
    LSL=v.getcell(cell='C19'), USL=v.getcell(cell='D19'))

v.run_mc()
v.plot('Gland Fill %')
v.forecasts['Gland Fill %'].getcapabilitymetrics()

