from vatic import *

v = vatic()
v.setnpts(1000)
# v.verbose = True

v.addassumption(cell='C8', tag=v.getcell(cell='B8'), 
    dist={N: (v.getcell(cell='C8'), v.getcell(cell='D8')/3)})
v.addassumption(cell='C9', tag=v.getcell(cell='B9'), 
    dist={N: (v.getcell(cell='C9'), v.getcell(cell='D9')/3)})
v.addassumption(cell='C10', tag=v.getcell(cell='B10'), 
    dist={N: (v.getcell(cell='C10'), v.getcell(cell='D10')/3)})
v.addassumption(cell='C11', tag=v.getcell(cell='B11'), 
    dist={N: (v.getcell(cell='C11'), v.getcell(cell='D11')/3)})
v.addassumption(cell='C12', tag=v.getcell(cell='B12'), 
    dist={N: (v.getcell(cell='C12'), v.getcell(cell='D12')/3)})
v.addassumption(cell='C13', tag=v.getcell(cell='B13'), 
    dist={N: (v.getcell(cell='C13'), v.getcell(cell='D13')/3)})
v.addassumption(cell='C14', tag=v.getcell(cell='B14'), 
    dist={N: (v.getcell(cell='C14'), v.getcell(cell='D14')/3)})
v.addassumption(cell='C15', tag=v.getcell(cell='B15'), 
    dist={N: (v.getcell(cell='C15'), v.getcell(cell='D15')/3)})

v.addforecast(cell='C24', tag=v.getcell(cell='B24'))
v.addforecast(cell='C25', tag=v.getcell(cell='B25'))
v.addforecast(cell='C26', tag=v.getcell(cell='B26'))
v.addforecast(cell='C30', tag=v.getcell(cell='B30'), 
    LSL=v.getcell(cell='C19'), USL=v.getcell(cell='D19'))
v.addforecast(cell='C31', tag=v.getcell(cell='B31'), 
    LSL=v.getcell(cell='C20'), USL=v.getcell(cell='D20'))

v.run_mc()
v.plot('Seal Comp. %')
v.plot('Gland Fill %')
sccap = v.forecasts['Seal Comp. %'].getcapabilitymetrics()
gfcap = v.forecasts['Gland Fill %'].getcapabilitymetrics()

print('Seal Comp. % Capability Metrics:\n', sccap)
print('Gland Fill % Capability Metrics:\n', gfcap)


