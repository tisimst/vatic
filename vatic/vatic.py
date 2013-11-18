"""
Vatic: open source risk analysis

Vatic is designed to be a viable open-source alternative to commercial
packages like Crystall Ball, @Risk, and ModelRisk. It uses Monte Carlo
simulation to give clarity to activities that have uncertianty, reducing
risk, predicting success, and improving the bottom line.

Requirements
------------

To be able to run Vatic, the following additional Python packages must be 
installed:

- win32com
- NumPy
- SciPy
- Matplotlib



"""

from win32com.client import Dispatch
import numpy as np
import matplotlib.pyplot as plt
import mcerp
from mcerp import *
from time import time

class vatic:

    def __init__(self):
        self.inputs = {}
        self.outputs = {}
        self.iter = 0
        self.xl = Dispatch('Excel.Application')
        info = """
Welcome to Vatic! For your information, to speed up processing, you should 
remove any plots or any unnecessary calculations. Hope it's useful!
"""
        print info

    def setnpts(self, n):
        mcerp.npts = n
        
    def getcell(self, cellref=None):
        if cellref is None:
            val = self.xl.ActiveSelection.Value
        else:
            val = self.xl.Range(cellref).Value
        return val
            
    def __repr__(self):
        intags = self.inputs.keys()
        outtags = self.outputs.keys()
        tmp = 'VATIC: OPEN SOURCE RISK ANALYSIS\n'
        tmp += '*'*65
        tmp += '\n'
        tmp += 'Model Inputs ({}):\n'.format(len(intags))
        tmp += ('='*65)+'\n'
        for tag in intags:
            tmp += str(self.inputs[tag])+'\n'
        tmp += '\nModel Outputs ({}):\n'.format(len(outtags))
        tmp += ('='*65)+'\n'
        for tag in outtags:
            tmp += str(self.outputs[tag])+'\n'
        return tmp
    
    def __str__(self):
        return repr(self)
        
    def addinput(self, cellref=None, tag=None, dist=None):
        if cellref is None:
            cellref = self.xl.Selection.Address
        newvar = InputVariable(cellref, tag, dist)
        self.inputs[tag] = newvar
        self.xl.Range(cellref).Interior.ColorIndex = 4  # 4 = Green
        print 'Added input "{}": {}'.format(tag, self.xl.Range(cellref).Value)
    
    def addoutput(self, cellref=None, tag=None, LSL=None, USL=None):
        if cellref is None:
            cellref = self.xl.Selection.Address
        newvar = OutputVariable(cellref, tag, LSL, USL)
        self.outputs[tag] = newvar
        self.xl.Range(cellref).Interior.ColorIndex = 8  # 8 = Cyan
        print 'Added output "{}": {}'.format(tag, self.xl.Range(cellref).Value)
        
    def run_mc(self):
        # initialize the result vector
        npts = mcerp.npts
        res = np.zeros((npts, len(self.outputs)))  # npts comes from mcerp
        intags = self.inputs.keys()
        outtags = self.outputs.keys()
        
        # save the original input values (for later re-setting)
        start_vals = {}
        for tag in intags:
            var = self.inputs[tag]
            start_vals[tag] = self.xl.Range(var.cellref).Value
        
        # reset the output samples
        for tag in outtags:
            self.outputs[tag].samples = np.zeros(mcerp.npts)
        
        try:
            # run the simulations
            print 'Simulating Now!'
            print 'Running {} iterations...'.format(mcerp.npts)
            start = time()
            in_matrix = np.zeros((mcerp.npts, len(intags)))
            in_cellrefs = []
            out_cellrefs = []
            for i, tag in enumerate(intags):
                in_matrix[:, i] = self.inputs[tag].dist._mcpts[:]
                in_cellrefs.append(self.inputs[tag].cellref)
                
            for i, tag in enumerate(outtags):
                out_cellrefs.append(self.outputs[tag].cellref)
                
            for it in xrange(npts):
                if it%(npts/20)==0 and it!=0:
                    timetemplate = '{} of {} complete (est. {:2.1f} minutes left)'
                    timeperiter = (time()-start)/(it + 1)
                    print timetemplate.format(it, npts, timeperiter*(npts - it)/60)
                    print 'Current Data:'
                    print ' '
                    print '         Output              Mean         Stdev        Skewness      Kurtosis'
                    print ' ----------------------  ------------  ------------  ------------  ------------'
                    stattemplate = ' %22s  %12.4f  %12.4f  %12.4f  %12.4f'
                    stats = self.get_output_stats(upto=self.iter)
                    for i, tag in enumerate(outtags):
                        mn, vr, sk, kt, dmin, dmax = stats[:, i]
                        sd = np.sqrt(vr)
                        print stattemplate%(tag[:min([len(tag), 22])], mn, sd, sk, kt)
                    print ' '
                    
                # freeze the calculations until all the inputs have been changed
                self.xl.ActiveSheet.EnableCalculation = 0
                
                # loop through the inputs and put in a new sampled value
                for j in xrange(len(in_cellrefs)):
                    self.xl.Range(in_cellrefs[j]).Value = in_matrix[it, j]
                
                # manually calculate to see what has been changed
                self.xl.ActiveSheet.EnableCalculation = 1
                self.xl.Calculate()

                # loop through the outputs and get the newly calculated value
                for j, tag in enumerate(outtags):
                    tmp = self.xl.Range(out_cellrefs[j]).Value
                    self.outputs[tag].samples[it] = tmp
                    
                # track the number of iterations
                self.iter = it + 1
                
        except BaseException:
            print 'An error occured. Resetting sheet to original values.'
            raise
        
        finally:
            # re-set the original input values
            for tag in intags:
                var = self.inputs[tag]
                self.xl.Range(var.cellref).Value = start_vals[tag]
            for tag in outtags:
                var = self.outputs[tag]
                var.samples = var.samples[:self.iter]
            
            # reset to automatically calculate
            self.xl.ActiveSheet.EnableCalculation = 1
            
            print 'SIMULATIONS COMPLETE! (Took approximately {:2.1f} minutes)'.format((time()-start)/60)
    
    def plot(self, tag=None):
        if tag is None:  # plot all outputs
            outtags = self.outputs.keys()
            for tag in outtags:
                plt.figure()
                var = self.outputs[tag]
                tmp = mcerp.UncertainFunction(var.samples[:self.iter])
                tmp.plot(hist=True)
                plt.title('Histogram of '+tag)
            plt.show()
        elif tag in self.outputs.keys():
            var = self.outputs[tag]
            tmp = mcerp.UncertainFunction(var.samples[:self.iter])
            tmp.plot(hist=True)
            plt.title('Histogram of '+tag)
            plt.show()
        elif tag in self.inputs.keys():
            var = self.inputs[tag]
            var.dist.plot(hist=True)
            plt.title('Histogram of '+tag)
            plt.show()
    
    def reset(self):
        self.inputs = {}
        self.outputs = {}
        self.samples = []
        
    def get_input_stats(self, upto=None):
        """
        Calculate some basic statistics for the inputs:
        1. Mean
        2. Variance
        3. Standardized Skewness Coefficient
        4. Standardized Kurtosis Coefficient
        5. Minimum
        6. Maximum
        
        Optional
        --------
        upto : int
            The maximum integer location that should be included in the
            calculations.
        """
        intags = self.inputs.keys()
        stats = np.zeros((6, len(intags)))
        for tag in intags:
            stats[:, i] = self.inputs[tag].getstats(upto)
        return stats

    def get_output_stats(self, upto=None):
        """
        Calculate some basic statistics for the outputs:
        1. Mean
        2. Variance
        3. Standardized Skewness Coefficient
        4. Standardized Kurtosis Coefficient
        5. Minimum
        6. Maximum
        
        Optional
        --------
        upto : int
            The maximum integer location that should be included in the
            calculations.
        """
        outtags = self.outputs.keys()
        stats = np.zeros((6, len(outtags)))
        for i, tag in enumerate(outtags):
            stats[:, i] = self.outputs[tag].getstats(upto)
        return stats

class InputVariable:
    def __init__(self, cellref=None, tag=None, dist=None):
        self.cellref = cellref
        self.tag = tag
        self.dist = dist
    
    def __repr__(self):
        template = '{}:\n    cellref = {}\n    {} samples'
        return template.format(self.tag, self.cellref, len(self.dist._mcpts))
        
    def __str__(self):
        return repr(self)
        
    def getstats(self, upto=None):
        """
        Calculate some basic statistics for the samples:
        1. Mean
        2. Variance
        3. Standardized Skewness Coefficient
        4. Standardized Kurtosis Coefficient
        5. Minimum
        6. Maximum
        
        Optional
        --------
        upto : int
            The maximum integer location that should be included in the
            calculations.
        """
        if upto is None:
            data = self.dist._mcpts[:]
        else:
            data = self.dist._mcpts[:upto]
        mn = np.mean(data)
        vr = np.mean((data - mn)**2)
        sd = np.sqrt(vr)
        sk = 0.0 if abs(sd)<=1e-8 else np.mean((data - mn)**3)/sd**3
        kt = 0.0 if abs(sd)<=1e-8 else np.mean((data - mn)**4)/sd**4
        dmin = np.min(data)
        dmax = np.max(data)
        stats = np.array([mn, vr, sk, kt, dmin, dmax])
        return stats

class OutputVariable:
    def __init__(self, cellref=None, tag=None, LSL=None, USL=None):
        self.cellref = cellref
        self.tag = tag
        self.LSL = LSL
        self.USL = USL
        self.samples = []

    def __repr__(self):
        template = '{}:\n    cellref = {}\n    LSL = {}\n    USL = {}\n    {} samples'
        return template.format(self.tag, self.cellref, self.LSL, self.USL, len(self.samples))
        
    def __str__(self):
        return repr(self)
        
    def getstats(self, upto=None):
        """
        Calculate some basic statistics for the samples:
        1. Mean
        2. Variance
        3. Standardized Skewness Coefficient
        4. Standardized Kurtosis Coefficient
        5. Minimum
        6. Maximum
        
        Optional
        --------
        upto : int
            The maximum integer location that should be included in the
            calculations.
        """
        if upto is None:
            data = self.samples[:]
        else:
            data = self.samples[:upto]
        mn = np.mean(data)
        vr = np.mean((data - mn)**2)
        sd = np.sqrt(vr)
        sk = 0.0 if abs(sd)<=1e-8 else np.mean((data - mn)**3)/sd**3
        kt = 0.0 if abs(sd)<=1e-8 else np.mean((data - mn)**4)/sd**4
        dmin = np.min(data)
        dmax = np.max(data)
        stats = np.array([mn, vr, sk, kt, dmin, dmax])
        return stats
