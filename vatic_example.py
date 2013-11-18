from vatic import *

v = vatic()
v.setnpts(10000)

v.addinput(cellref='C8', tag=v.getcell('B8'), dist=N(v.getcell('C8'), v.getcell('D8')/3))
v.addinput(cellref='C9', tag=v.getcell('B9'), dist=N(v.getcell('C9'), v.getcell('D9')/3))
v.addinput(cellref='C10', tag=v.getcell('B10'), dist=N(v.getcell('C10'), v.getcell('D10')/3))
v.addinput(cellref='C11', tag=v.getcell('B11'), dist=N(v.getcell('C11'), v.getcell('D11')/3))
v.addinput(cellref='C12', tag=v.getcell('B12'), dist=N(v.getcell('C12'), v.getcell('D12')/3))
v.addinput(cellref='C13', tag=v.getcell('B13'), dist=N(v.getcell('C13'), v.getcell('D13')/3))
v.addinput(cellref='C14', tag=v.getcell('B14'), dist=N(v.getcell('C14'), v.getcell('D14')/3))
v.addinput(cellref='C15', tag=v.getcell('B15'), dist=N(v.getcell('C15'), v.getcell('D15')/3))

v.addoutput(cellref='C24', tag=v.getcell('B24'))
v.addoutput(cellref='C25', tag=v.getcell('B25'))
v.addoutput(cellref='C26', tag=v.getcell('B26'))
v.addoutput(cellref='C30', tag=v.getcell('B30'), LSL=v.getcell('C19'), USL=v.getcell('D19'))
v.addoutput(cellref='C31', tag=v.getcell('B31'), LSL=v.getcell('C20'), USL=v.getcell('D20'))

#print v
v.run_mc()
v.plot('Seal Comp. %')
v.plot('Gland Fill %')
