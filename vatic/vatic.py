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
Welcome to Vatic! FYI: to speed up processing, you should remove
any plots or any unnecessary calculations. Hope it's useful!
"""
        print info

    def setnpts(self, n):
        mcerp.npts = n
        
    def getcell(self, cellref=None, workbook=None, worksheet=None):
        if workbook is None:
            workbook = self.xl.ActiveWorkbook
        if worksheet is None:
            worksheet = self.xl.ActiveSheet
        if cellref is None:
            cellref = self.xl.Selection.Address
        val = self.xl.Workbook(workbook).Worksheet(worksheet).Range(cellref).Value
        return val
            
    def __repr__(self):
        assumtags = self.assumptions.keys()
        decistags = self.decisions.keys()
        foretags = self.forecasts.keys()
        tmp = 'VATIC: OPEN SOURCE RISK ANALYSIS\n'
        tmp += '*'*65
        tmp += '\n'
        tmp += 'Model ASSUMPTION variables ({}):\n'.format(len(assumtags))
        tmp += ('='*65)+'\n'
        for tag in assumtags:
            tmp += str(self.assumptions[tag])+'\n'
        tmp += '\n'
        tmp += 'Model DECISION variables ({}):\n'.format(len(decistags))
        tmp += ('='*65)+'\n'
        for tag in decistags:
            tmp += str(self.decisions[tag])+'\n'
        tmp += '\nModel FORECAST variables ({}):\n'.format(len(foretags))
        tmp += ('='*65)+'\n'
        for tag in foretags:
            tmp += str(self.forecasts[tag])+'\n'
        return tmp
    
    def __str__(self):
        return repr(self)
        
    def newassumption(self, cellref=None, tag=None, dist=None, workbook=None, 
    worksheet=None):
        """
        Adds a new ASSUMPTION variable (i.e., a contributing input that is
        more susceptible to variability, such as marketing costs).
        
        Parameters
        ----------
        cellref : str
            The Excel cell callout for the input (e.g., 'A3'). If not given, 
            it is assumed that the cellref comes from the current selection.
        tag : str
            The name this cell will be referred to (e.g., 'Marketing costs').
            If none is given, then the cellref becomes the tag. 
        dist : dict
            The statistical distribution information, in the form 
            "{DIST: (PARAM1, [PARAM2, ...])}"
        workbook : str
            The Excel workbook file name that the input cell comes from (e.g.,
            'costanalysis.xlx'). If not given, it is assumed that the cellref
            comes from the active workbook.
        worksheet : str
            The Excel sheet name that the input cell comes from (e.g., 
            'Sheet1'). If not given, it is assumed that the cellref
            comes from the active sheet.

        Examples
        --------
        
        Create an assumption from the active selection (cell B3) and assign
        a Normal distribution with a mean of 24 and standard deviation of 1::
        
            >>> from vatic import *
            >>> v = vatic()
            >>> v.newassumption(cellref='B3', dist={N: (24, 1)})
            
        """
        if workbook is None:
            workbook = self.xl.ActiveWorkbook
        if worksheet is None:
            worksheet = self.xl.ActiveSheet
        if cellref is None:
            cellref = self.xl.Selection.Address
        newvar = AssumptionVariable(cellref, tag, dist, workbook, worksheet)
        self.assumptions[tag] = newvar
        self.xl.Range(cellref).Interior.ColorIndex = 4  # 4 = Green
        print 'Added ASSUMPTION variable "{}": {}'.format(tag, 
            self.xl.Workbook(workbook).Worksheet(worksheet).Range(cellref).Value)
    
    def newdecision(self, cellref=None, tag=None, options=None, workbook=None, 
        worksheet=None):
        """
        Adds a new DECISION variable (i.e., a contributing input that you
        would have more control over, like choosing to advertise product A 
        or not).
        
        Parameters
        ----------
        cellref : str
            The Excel cell callout for the input (e.g., 'B4'). If not given, 
            it is assumed that the cellref comes from the current selection.
        tag : str
            The name this cell will be referred to (e.g., 'Marketing costs').
            If none is given, then the cellref becomes the tag. 
        choices : array-like
            The allowable values that you are able to choose from (e.g.
            ['yes', 'no']).
        workbook : str
            The Excel workbook file name that the input cell comes from (e.g.,
            'costanalysis.xlx'). If not given, it is assumed that the cellref
            comes from the active workbook.
        worksheet : str
            The Excel sheet name that the input cell comes from (e.g., 
            'Sheet1'). If not given, it is assumed that the cellref
            comes from the active sheet.

        Examples
        --------
        
        Create an decision from the active selection (cell B4) and assign
        a list of options::
        
            >>> from vatic import *
            >>> v = vatic()
            >>> v.newdecision(cellref='B4', tag='Choose product A', 
            ...     options=['yes', 'no'])
            
        """
        if workbook is None:
            workbook = self.xl.ActiveWorkbook
        if worksheet is None:
            worksheet = self.xl.ActiveSheet
        if cellref is None:
            cellref = self.xl.Selection.Address
        newvar = DecisionVariable(cellref, tag, options, workbook, worksheet)
        self.decisions[tag] = newvar
        self.xl.Range(cellref).Interior.ColorIndex = 6  # 6 = Yellow
        print 'Added DECISION variable "{}": {}'.format(tag, 
            self.xl.Workbook(workbook).Worksheet(worksheet).Range(cellref).Value)
    
    def newforecast(self, cellref=None, tag=None, LSL=None, USL=None,
        target=None, workbook=None, worksheet=None):
        """
        Adds a new FORECAST variable (i.e., a variable that is calculated
        rather than specified, like net profits).
        
        Parameters
        ----------
        cellref : str
            The Excel cell callout for the input (e.g., 'B4'). If not given, 
            it is assumed that the cellref comes from the current selection.
        tag : str
            The name this cell will be referred to (e.g., 'Marketing costs').
            If none is given, then the cellref becomes the tag. 
        LSL : scalar
            The 'Lower-Specification-Limit' of the forecast which designates
            the minimum allowable value. This is utilized when calculating 
            various statistics and capability metrics. If none given, then
            LSL = -Infinity.
        USL : scalar
            The 'Upper-Specification-Limit' of the forecast which designates
            the maximum allowable value. This is utilized when calculating 
            various statistics and capability metrics. If none given, then
            USL = +Infinity.
        target : scalar
            The target-value of the forecast. This is utilized when 
            calculating various statistics and capability metrics. No default
            is given for this parameter.
        workbook : str
            The Excel workbook file name that the input cell comes from (e.g.,
            'costanalysis.xlx'). If not given, it is assumed that the cellref
            comes from the active workbook.
        worksheet : str
            The Excel sheet name that the input cell comes from (e.g., 
            'Sheet1'). If not given, it is assumed that the cellref
            comes from the active sheet.

        Examples
        --------
        
        Create an forecast from the active selection (cell B5) and assigns
        LSL value of 0 for showing certainty of making a profit::
        
            >>> from vatic import *
            >>> v = vatic()
            >>> v.newforecast(cellref='B5', tag='Net Profits', LSL=0)
            
        """
        if workbook is None:
            workbook = self.xl.ActiveWorkbook
        if worksheet is None:
            worksheet = self.xl.ActiveSheet
        if cellref is None:
            cellref = self.xl.Selection.Address
        newvar = OutputVariable(cellref, tag, LSL, USL, target, workbook, 
            worksheet)
        self.forecasts[tag] = newvar
        self.xl.Range(cellref).Interior.ColorIndex = 8  # 8 = Cyan
        print 'Added FORECAST variable "{}": {}'.format(tag, 
            self.xl.Workbook(workbook).Worksheet(worksheet).Range(cellref).Value)
        
    def run_mc(self):
        # Disable charts and stuff
        self.xl.ScreenUpdating = 0

        # initialize the result vector
        npts = mcerp.npts
        res = np.zeros((npts, len(self.forecasts)))  # npts comes from mcerp
        assumtags = self.assumptions.keys()
        decistags = self.decisions.keys()
        foretags = self.forecasts.keys()
        
        # save the original input values (for later re-setting)
        start_vals = {}
        for tag in assumtags:
            var = self.assumptions[tag]
            start_vals[tag] = self.getcell(var.cellref, var.workbook,
                var.worksheet).Value
        
        # reset the output samples
        for tag in foretags:
            self.forecasts[tag].samples = np.zeros(npts)
        
        try:
            # run the simulations
            print 'Simulating Now!'
            print 'Running {} iterations...'.format(npts)
            start = time()
            a_matrix = np.zeros((npts, len(assumtags)))
            a_cellrefs = []
            f_cellrefs = []
            for i, tag in enumerate(assumtags):
                a_matrix[:, i] = self.assumptions[tag].dist._mcpts[:]
                a_cellrefs.append(self.assumptions[tag].cellref)
                
            for i, tag in enumerate(foretags):
                f_cellrefs.append(self.forecasts[tag].cellref)
                
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
                    for i, tag in enumerate(foretags):
                        mn, vr, sk, kt, dmin, dmax = stats[:, i]
                        sd = np.sqrt(vr)
                        print stattemplate%(tag[:min([len(tag), 22])], mn, sd, sk, kt)
                    print ' '
                    
                # freeze the calculations until all the inputs have been changed
                self.xl.ActiveSheet.EnableCalculation = 0
                
                # loop through the inputs and put in a new sampled value
                for j in xrange(len(a_cellrefs)):
                    self.xl.Range(a_cellrefs[j]).Value = a_matrix[it, j]
                
                # manually calculate to see what has been changed
                self.xl.ActiveSheet.EnableCalculation = 1
                self.xl.Calculate()

                # loop through the outputs and get the newly calculated value
                for j, tag in enumerate(foretags):
                    tmp = self.xl.Range(f_cellrefs[j]).Value
                    self.forecasts[tag].samples[it] = tmp
                    
                # track the number of iterations
                self.iter = it + 1
                
        except BaseException:
            print 'An error occured. Resetting sheet to original values.'
            raise
        
        finally:
            # re-set the original input values
            for tag in assumtags:
                var = self.assumptions[tag]
                self.xl.Range(var.cellref).Value = start_vals[tag]
            for tag in foretags:
                var = self.forecasts[tag]
                var.samples = var.samples[:self.iter]
            
            # reset to automatically calculate
            self.xl.ActiveSheet.EnableCalculation = 1
            
            # Re-enable charts and stuff
            self.xl.ScreenUpdating = 1

            print 'SIMULATIONS COMPLETE! (Took approximately {:2.1f} minutes)'.format((time()-start)/60)
    
    def plot(self, tag=None):
        if tag is None:  # plot all forecasts
            foretags = self.forecasts.keys()
            for tag in foretags:
                plt.figure()
                var = self.forecasts[tag]
                tmp = mcerp.UncertainFunction(var.samples[:self.iter])
                tmp.plot(hist=True)
                plt.title('Histogram of '+tag)
        elif tag in self.forecasts.keys():
            var = self.forecasts[tag]
            tmp = mcerp.UncertainFunction(var.samples[:self.iter])
            tmp.plot(hist=True)
            plt.title('Histogram of '+tag)
        elif tag in self.assumptions.keys():
            var = self.assumptions[tag]
            var.dist.plot(hist=True)
            plt.title('Histogram of '+tag)
        plt.show()
    
    def reset(self):
        self.assumptions = {}
        self.decisions = {}
        self.forecasts = {}
        self.samples = []
        
    def get_assumption_stats(self, upto=None):
        """
        Calculate some basic statistics for the assumptions:
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
        assumtags = self.assumptions.keys()
        stats = np.zeros((6, len(assumtags)))
        for tag in assumtags:
            stats[:, i] = self.assumptions[tag].getstats(upto)
        return stats

    def get_forecast_stats(self, upto=None):
        """
        Calculate some basic statistics for the forecasts:
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
        foretags = self.forecasts.keys()
        stats = np.zeros((6, len(foretags)))
        for i, tag in enumerate(foretags):
            stats[:, i] = self.forecasts[tag].getstats(upto)
        return stats

class AssumptionVariable:
    def __init__(self, cellref=None, tag=None, dist=None, workbook=None,
        worksheet=None):
        self.cellref = cellref
        self.tag = tag
        self.dist = dist
        self.workbook = workbook
        self.worksheet = worksheet
    
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
        sd = (vr)**0.5
        sk = 0.0 if abs(sd)<=1e-8 else np.mean((data - mn)**3)/sd**3
        kt = 0.0 if abs(sd)<=1e-8 else np.mean((data - mn)**4)/sd**4
        dmin = np.min(data)
        dmax = np.max(data)
        stats = np.array([mn, vr, sk, kt, dmin, dmax])
        return stats

class DecisionVariable:
    def __init__(self, cellref=None, tag=None, options=None, workbook=None,
        worksheet=None):
        self.cellref = cellref
        self.tag = tag
        self.options = options
        self.workbook = workbook
        self.worksheet = worksheet
    
    def __repr__(self):
        template = '{}:\n    cellref = {}\n    options = {}'
        return template.format(self.tag, self.cellref, self.options)
        
    def __str__(self):
        return repr(self)
        
class ForecastVariable:
    def __init__(self, cellref=None, tag=None, LSL=None, USL=None, target=None,
        workbook=None, worksheet=None):
        self.cellref = cellref
        self.tag = tag
        self.LSL = LSL
        self.USL = USL
        self.target = target
        self.samples = []
        self.workbook = workbook
        self.worksheet = worksheet

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
        

