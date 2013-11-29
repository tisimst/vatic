"""
vatic: Open Source Risk Analysis

This package is designed for those who cannot afford commercial risk analysis
programs that interface with Microsoft Excel.

vatic requires the following packages:
- win32com
- numpy
- matplotlib
- scipy
- mcerp

To use vatic, you 1) select cells whose value you deem to be uncertain and
identify them as an ASSUMPTION cell and assign a statistical distribution to
them that gives the a range of possible values. Then, 2) select cells that are
the result of some calculation that uses the ASSUMPTION cells. These output
cells are identified as FORECAST cells. 3) Once both kinds of cells have been
identified, running a MONTE CARLO SIMULATION will show you how the uncertainty
in the ASSUMPTION cells is propagated to the FORECAST cells, so that you can
4) ANALYZE the results to make predictions based on the variable outputs.
"""
from __future__ import division, print_function
from win32com.client import Dispatch
import numpy as np
import matplotlib.pyplot as plt
import mcerp
from mcerp import *
from time import time

class vatic:

    def __init__(self):
        self.assumptions = {}
        self.forecasts = {}
        self.iter = 0
        self.xl = Dispatch('Excel.Application')
        info = """
Welcome to Vatic! FYI: to speed up processing, you should remove
any plots or any unnecessary calculations. Hope it's useful!
"""
        print(info)
        self.verbose = False

    def setnpts(self, n):
        mcerp.npts = n
        
    def getcell(self, vobject=None, cell=None, ws=None, wb=None):
        """
        Get the displayed value from a cell. There are two main options here,
        either by inputing an ASSUMPTION or FORECAST object using the kwarg
        ``vobject`` (from which the target cell reference should already 
        exist), or by explicitly identifying the cell-reference (kwarg 
        ``cell``), worksheet (kwarg ``ws``), and workbook (kwarg ``wb``).
        """
        # print('getcell\n-------')
        if vobject is not None:
            if vobject.tag in self.assumptions or vobject.tag in self.forecasts:
                # print('    vobject: {}'.format(vobject))
                wb = vobject.wb
                ws = vobject.ws
                cell = vobject.cell
        if wb is None:
            wb = str(self.xl.ActiveWorkbook.Name)
        if ws is None:
            ws = str(self.xl.ActiveSheet.Name)
        if cell is None:
            cell = self.xl.Selection.Address
        # print('    wb: {}'.format(wb))
        # print('    ws: {}'.format(ws))
        # print('    cell: {}'.format(cell))
        val = self.xl.Workbooks(wb).Worksheets(ws).Range(cell).Value
        # print('    val: {}'.format(val))
        return val
            
    def setcell(self, vobject=None, value=None, cell=None, ws=None, 
        wb=None):
        """
        Set the displayed value of a cell (kwarg ``value``). There are two 
        main options here, either by inputing an ASSUMPTION or FORECAST object
        using the kwarg ``vobject`` (from which the target cell reference 
        should already exist), or by explicitly identifying the cell-
        reference (kwarg ``cell``), worksheet (kwarg ``ws``), and workbook 
        (kwarg ``wb``).
        """
        if vobject is not None:
            if vobject.tag in self.assumptions or vobject.tag in self.forecasts:
                wb = vobject.wb
                ws = vobject.ws
                cell = vobject.cell
        if wb is None:
            wb = str(self.xl.ActiveWorkbook.Name)
        if ws is None:
            ws = str(self.xl.ActiveSheet.Name)
        if cell is None:
            cell = str(self.xl.Selection.Address)
        self.xl.Workbooks(wb).Worksheets(ws).Range(cell).Value = value
    
    def setcellbackground(self, vobject=None, color=None, cell=None, ws=None,
        wb=None):
        """
        Set a cell's background color. Valid values for kwarg ``color`` are:
        - Black: "k" or 1
        - White: "w" or 2
        - Red: "r" or 3
        - Green: "g" or 4
        - Blue: "b" or 5
        - Yellow: "y" or 6
        - Magenta: "m" or 7
        - Cyan: "c" or 8
        - Some other index 9-56 (see dmicritchie.mvps.org/excel/colors.htm for
          more details of what colors these indicies correspond to)
        """
        if vobject is not None:
            if vobject.tag in self.assumptions or vobject.tag in self.forecasts:
                wb = vobject.wb
                ws = vobject.ws
                cell = vobject.cell
        if wb is None:
            wb = str(self.xl.ActiveWorkbook.Name)
        if ws is None:
            ws = str(self.xl.ActiveSheet.Name)
        if cell is None:
            cell = str(self.xl.Selection.Address)
        target = self.xl.Workbooks(wb).Worksheets(ws).Range(cell)

        validcolornames = {'k': 1, 'w': 2, 
                           'r': 3, 'g': 4, 'b': 5, 
                           'y': 6, 'm': 7, 'c': 8}
                       
        if isinstance(color, str):
            color = color.lower()
            try:
                color = validcolornames[color]
            except:
                raise
        
        color = int(color)
        assert 1<=color<=56, 'Color index must be a value between 1 and 56 or an equivalent letter (see doc)'
        target.Interior.ColorIndex = color
        
    def setforebackground(self, vobject=None, color=None, cell=None, ws=None,
        wb=None):
        """
        Set a cell's foreground (font) color. Valid values for kwarg 
        ``color`` are:
        - Black: "k" or 1
        - White: "w" or 2
        - Red: "r" or 3
        - Green: "g" or 4
        - Blue: "b" or 5
        - Yellow: "y" or 6
        - Magenta: "m" or 7
        - Cyan: "c" or 8
        - Some other index 9-56 (see dmicritchie.mvps.org/excel/colors.htm for
          more details of what colors these indicies correspond to)
        """
        if vobject is not None:
            if vobject.tag in self.assumptions or vobject.tag in self.forecasts:
                wb = vobject.wb
                ws = vobject.ws
                cell = vobject.cell
        if wb is None:
            wb = str(self.xl.ActiveWorkbook.Name)
        if ws is None:
            ws = str(self.xl.ActiveSheet.Name)
        if cell is None:
            cell = str(self.xl.Selection.Address)
        target = self.xl.Workbooks(wb).Worksheets(ws).Range(cell)

        validcolornames = {'k': 1, 'w': 2, 
                           'r': 3, 'g': 4, 'b': 5, 
                           'y': 6, 'm': 7, 'c': 8}
                       
        if isinstance(color, str):
            color = color.lower()
            try:
                color = validcolornames[color]
            except:
                raise
                
        color = int(color)
        assert 1<=color<=56, 'Color index must be a value between 1 and 56 or an equivalent letter (see doc)'
        target.Font.ColorIndex = color

    def __repr__(self):
        atags = self.assumptions.keys()
        ftags = self.forecasts.keys()
        tmp = 'VATIC: OPEN SOURCE RISK ANALYSIS\n'
        tmp += '*'*65
        tmp += '\n'
        tmp += 'Model ASSUMPTION variables ({}):\n'.format(len(atags))
        tmp += ('='*65)+'\n'
        for tag in atags:
            tmp += str(self.assumptions[tag])+'\n'
        tmp += '\nModel FORECAST variables ({}):\n'.format(len(ftags))
        tmp += ('='*65)+'\n'
        for tag in ftags:
            tmp += str(self.forecasts[tag])+'\n'
        return tmp
    
    def __str__(self):
        return repr(self)
        
    def addassumption(self, cell=None, tag=None, dist=None, wb=None, 
        ws=None):
        """
        Adds a new ASSUMPTION variable (i.e., a contributing input that is
        more susceptible to variability, such as marketing costs).
        
        Parameters
        ----------
        cell : str
            The Excel cell callout for the input (e.g., 'A3'). If not given, 
            it is assumed that the cell comes from the current selection.
        tag : str
            The name this cell will be referred to (e.g., 'Marketing costs').
            If none is given, then the cell becomes the tag. 
        dist : dict
            The statistical distribution information, in the form 
            "{DIST: (PARAM1, [PARAM2, ...])}"
        wb : str
            The Excel workbook file name that the input cell comes from (e.g.,
            'costanalysis.xlx'). If not given, it is assumed that the cell
            comes from the active workbook.
        ws : str
            The Excel sheet name that the input cell comes from (e.g., 
            'Sheet1'). If not given, it is assumed that the cell
            comes from the active sheet.

        Examples
        --------
        
        Create an assumption from the active selection (cell B3) and assign
        a Normal distribution with a mean of 24 and standard deviation of 1::
        
            >>> from vatic import *
            >>> v = vatic()
            >>> v.addassumption(cell='B3', dist={N: (24, 1)})
            
        """
        if wb is None:
            wb = str(self.xl.ActiveWorkbook.Name)
        if ws is None:
            ws = str(self.xl.ActiveSheet.Name)
        if cell is None:
            cell = str(self.xl.Selection.Address)
        newvar = AssumptionVariable(cell, tag, dist, wb, ws)
        self.assumptions[tag] = newvar
        self.setcellbackground(newvar, 'g')
        if self.verbose:
            print('Added ASSUMPTION variable "{}": {}'.format(tag, 
                self.getcell(newvar)))
    
    def addforecast(self, cell=None, tag=None, LSL=None, USL=None,
        target=None, wb=None, ws=None):
        """
        Adds a new FORECAST variable (i.e., a variable that is calculated
        rather than specified, like net profits).
        
        Parameters
        ----------
        cell : str
            The Excel cell callout for the input (e.g., 'B4'). If not given, 
            it is assumed that the cell comes from the current selection.
        tag : str
            The name this cell will be referred to (e.g., 'Marketing costs').
            If none is given, then the cell becomes the tag. 
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
        wb : str
            The Excel workbook file name that the input cell comes from (e.g.,
            'costanalysis.xlx'). If not given, it is assumed that the cell
            comes from the active workbook.
        ws : str
            The Excel sheet name that the input cell comes from (e.g., 
            'Sheet1'). If not given, it is assumed that the cell
            comes from the active sheet.

        Examples
        --------
        
        Create an forecast from the active selection (cell B5) and assigns
        LSL value of 0 for showing certainty of making a profit::
        
            >>> from vatic import *
            >>> v = vatic()
            >>> v.addforecast(cell='B5', tag='Net Profits', LSL=0)
            
        """
        if wb is None:
            wb = str(self.xl.ActiveWorkbook.Name)
        if ws is None:
            ws = str(self.xl.ActiveSheet.Name)
        if cell is None:
            cell = str(self.xl.Selection.Address)
        newvar = ForecastVariable(cell, tag, LSL, USL, target, wb, ws)
        self.forecasts[tag] = newvar
        self.setcellbackground(newvar, 'c')
        if self.verbose:
            print('Added FORECAST variable "{}": {}'.format(tag, 
                self.getcell(newvar)))
        
    def run_mc(self):
        # Disable charts and stuff
        self.xl.ScreenUpdating = 0

        # initialize the result vector
        npts = mcerp.npts
        res = np.zeros((npts, len(self.forecasts)))  # npts comes from mcerp
        atags = self.assumptions.keys()
        ftags = self.forecasts.keys()
        
        # save the original input values (for later re-setting)
        start_vals = {}
        for tag in atags:
            var = self.assumptions[tag]
            start_vals[tag] = self.getcell(var)
        
        # reset the output samples
        for tag in ftags:
            self.forecasts[tag].samples = np.zeros(npts)
        
        try:
            # run the simulations
            print('Simulating Now!')
            print('Running {} iterations...'.format(npts))
            start = time()
            a_matrix = np.zeros((npts, len(atags)))
            a_cellrefs = []
            f_cellrefs = []
            for i, tag in enumerate(atags):
                avar = self.assumptions[tag]
                a_matrix[:, i] = avar.getnewsamples()
                
            for it in xrange(npts):
                if it%(npts/20)==0 and it!=0:
                    timetemplate = '{} of {} complete (est. {:2.1f} minutes left)'
                    timeperiter = (time()-start)/(it + 1)
                    print(timetemplate.format(it, npts, timeperiter*(npts - it)/60))
                    if self.verbose:
                        print('Current Data:')
                        print(' ')
                        print('         Output              Mean         Stdev        Skewness      Kurtosis')
                        print(' ----------------------  ------------  ------------  ------------  ------------')
                        stattemplate = ' %22s  %12.4f  %12.4f  %12.4f  %12.4f'
                        stats = self.get_forecast_stats(upto=self.iter)
                        for i, tag in enumerate(ftags):
                            mn, vr, sk, kt, dmin, dmax = stats[:, i]
                            sd = np.sqrt(vr)
                            print(stattemplate%(tag[:min([len(tag), 22])], mn, sd, sk, kt))
                        print(' ')
                    
                # freeze the calculations until all the inputs have been changed
                self.xl.ActiveSheet.EnableCalculation = 0
                
                # loop through the inputs and put in a new sampled value
                for j, tag in enumerate(atags):
                    self.setcell(self.assumptions[tag], a_matrix[it, j])
                
                # manually calculate to see what has been changed
                self.xl.ActiveSheet.EnableCalculation = 1
                self.xl.Calculate()

                # loop through the outputs and get the newly calculated value
                for tag in ftags:
                    tmp = self.getcell(self.forecasts[tag])
                    self.forecasts[tag].samples[it] = tmp
                    
                # track the number of iterations
                self.iter = it + 1
                
        except BaseException:
            print('An error occured. Resetting sheet to original values.')
            raise
        
        finally:
            # re-set the original input values
            for tag in atags:
                self.setcell(self.assumptions[tag], start_vals[tag])
            for tag in ftags:
                var = self.forecasts[tag]
                var.samples = var.samples[:self.iter]
            
            # reset to automatically calculate
            self.xl.ActiveSheet.EnableCalculation = 1
            
            # Re-enable charts and stuff
            self.xl.ScreenUpdating = 1

            if self.verbose:
                print('SIMULATIONS COMPLETE! (Took approximately {:2.1f} minutes)'.format((time()-start)/60))
    
    def plot(self, tag=None):
        if tag is None:  # plot all forecasts
            ftags = self.forecasts.keys()
            for tag in ftags:
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
            tmp = mcerp.UncertainFunction(var.samples[:self.iter])
            tmp.plot(hist=True)
            plt.title('Histogram of '+tag)
        plt.show()
    
    def reset(self):
        self.assumptions = {}
        self.forecasts = {}
        self.iter = 0
        
    def get_assumption_stats(self, assumption=None, upto=None):
        """
        Calculate some basic statistics for the assumptions:
        1. Mean
        2. Variance
        3. Standardized Skewness Coefficient
        4. Standardized Kurtosis Coefficient
        5. Minimum
        6. Maximum
        
        Parameters
        ----------
        assumption : str or AssumptionVariable
            The forecast of interest, either the actual object or its tag
            (default: None, in which all assumptions' stats will returned)
        upto : int
            The maximum integer location that should be included in the
            calculations.
        """
        if assumption is None:
            atags = self.assumptions.keys()
            stats = np.zeros((6, len(atags)))
            for tag in atags:
                stats[:, i] = self.assumptions[tag].getstats(upto)
            return stats
        elif isinstance(assumption, str) and assumption in self.assumptions:
            return self.assumptions[assumption].getstats(upto)
        elif isinstance(assumption, AssumptionVariable):
            return assumption.getstats(upto)

    def get_forecast_stats(self, forecast=None, upto=None):
        """
        Calculate some basic statistics for the forecasts:
        1. Mean
        2. Variance
        3. Standardized Skewness Coefficient
        4. Standardized Kurtosis Coefficient
        5. Minimum
        6. Maximum
        
        Parameters
        ----------
        forecast : str or ForecastVariable
            The forecast of interest, either the actual object or its tag
            (default: None, in which all forecasts' stats will returned)
        upto : int
            The maximum integer location that should be included in the
            calculations.
        """
        if forecast is None:
            ftags = self.forecasts.keys()
            stats = np.zeros((6, len(ftags)))
            for i, tag in enumerate(ftags):
                stats[:, i] = self.forecasts[tag].getstats(upto)
            return stats
        elif isinstance(forecast, str) and forecast in self.forecasts:
            return self.forecasts[forecast].getstats(upto)
        elif isinstance(forecast, ForecastVariable):
            return forecast.getstats(upto)

###############################################################################

class AssumptionVariable:
    def __init__(self, cell=None, tag=None, dist=None, wb=None, ws=None):
        self.cell = cell
        self.tag = tag
        self.dist = dist
        self.wb = wb
        self.ws = ws
        self.samples = None
        # print('AssumptionVariable:')
        # print('    cell: {}'.format(cell))
        # print('    tag: {}'.format(tag))
        # print('    dist: {}'.format(dist))
        # print('    wb: {}'.format(wb))
        # print('    ws: {}'.format(ws))
    
    def __repr__(self):
        tmp = ''
        if self.tag is not None:
            tmp += '{}:\n'.format(self.tag)
        if self.dist is not None:
            func = self.dist.keys()[0]
            params = self.dist[func]
            tmp += '    Distribution = {}{}\n'.format(func, params)
        if self.samples is not None:
            tmp += '    Number of samples = {}\n'.format(len(self.samples))
        if self.cell is not None:
            tmp += '    Cell = {}\n'.format(self.cell)
        if self.ws is not None:
            tmp += '    Worksheet = {}\n'.format(self.ws)
        if self.wb is not None:
            tmp += '    Workbook = {}\n'.format(self.wb)
        return tmp if tmp!='' else '<Empty AssumptionVariable>'
        
    def __str__(self):
        return repr(self)
        
    def getnewsamples(self):
        if self.dist is not None:
            func = self.dist.keys()[0]
            params = self.dist[func]
            self.samples = func(*params)._mcpts
            return self.samples
        else:
            print('Distribution not defined for {}'.format(self))
            return []
            
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
        if self.samples is not None:
            if upto is None:
                data = self.samples[:]
            else:
                data = self.samples[:upto]
            mn = np.mean(data)
            vr = np.mean((data - mn)**2)
            sd = (vr)**0.5
            sk = 0.0 if abs(sd)<=1e-8 else np.mean((data - mn)**3)/sd**3
            kt = 0.0 if abs(sd)<=1e-8 else np.mean((data - mn)**4)/sd**4
            dmin = np.min(data)
            dmax = np.max(data)
            stats = np.array([mn, vr, sk, kt, dmin, dmax])
            return stats

###############################################################################

class ForecastVariable:
    def __init__(self, cell=None, tag=None, LSL=None, USL=None, target=None,
        wb=None, ws=None):
        self.cell = cell
        self.tag = tag
        self.LSL = LSL
        self.USL = USL
        self.target = target
        self.samples = None
        self.wb = wb
        self.ws = ws

    def __repr__(self):
        tmp = ''
        if self.tag is not None:
            tmp += '{}:\n'.format(self.tag)
        if self.LSL is not None:
            tmp += '    LSL = {}\n'.format(self.LSL)
        if self.USL is not None:
            tmp += '    USL = {}\n'.format(self.USL)
        if self.target is not None:
            tmp += '    Target = {}\n'.format(self.target)
        if self.samples is not None:
            tmp += '    Number of samples = {}\n'.format(len(self.samples))
        if self.cell is not None:
            tmp += '    Cell = {}\n'.format(self.cell)
        if self.ws is not None:
            tmp += '    Worksheet = {}\n'.format(self.ws)
        if self.wb is not None:
            tmp += '    Workbook = {}\n'.format(self.wb)
        return tmp if tmp!='' else '<Empty ForecastVariable>'
        
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
        if self.samples is not None:
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
    
    def getcapabilitymetrics(self, zshift=1.5):
        """
        Calculate the capability metrics for a FORECAST variable. The
        individual metrics will only be calculated if the required FORECAST 
        criteria values have been identified (e.g., LSL, USL, and target).
        
        An optional input for ``zshift`` is also available (default: 1.5).
        
        The return value is a dictionary object with the capability metric
        name as the dict key and the metric's value as the dict value.
        
        Currently supported metrics include:
        
        - Cp, Cpk, Cpk-lower, Cpk-upper, Cpm (Short term)
        - Pp, Ppk, Ppk-lower, Ppk-upper, Ppm (Long term)
        - Z-LSL, Z-USL
        - Zst, Zst-total (Short term)
        - Zlt, Zlt-total (Long term)
        - p(N/C)-below, p(N/C)-above, p(N/C)-total
        - PPM-below, PPM-above, PPM-total
        """
        metrics = {}
        isLSL = self.LSL is not None
        isUSL = self.USL is not None
        istarget = self.target is not None
        LSL = self.LSL
        USL = self.USL
        target = self.target
        
        stats = self.getstats()
        mu = stats[0]
        sigma = stats[1]**0.5
        norm = ss.norm(loc=0, scale=1)
        normsdist = norm.pdf
        normsinv = norm.ppf
        
        if isLSL and isUSL:
            metrics['Cp'] = (USL - LSL)/(6*sigma)
        if isLSL:
            metrics['Cpk-lower'] = (mu - LSL)/(3*sigma)
        if isUSL:
            metrics['Cpk-upper'] = (USL - mu)/(3*sigma)
        if isLSL and isUSL:
            metrics['Cpk'] = min(metrics['Cpk-lower'], metrics['Cpk-upper'])
        if isLSL and isUSL and istarget:
            metrics['Cpm'] = (USL - LSL)/(6*((mu - target)^2 + sigma^2)**0.5)

        if isLSL and isUSL:
            metrics['Pp'] = (USL - LSL)/(6*sigma)
        if isLSL:
            metrics['Ppk-lower'] = (mu - LSL)/(3*sigma)
        if isUSL:
            metrics['Ppk-upper'] = (USL - mu)/(3*sigma)
        if isLSL and isUSL:
            metrics['Ppk'] = min(metrics['Ppk-lower'], metrics['Ppk-upper'])
        if isLSL and isUSL and istarget:
            metrics['Ppm'] = (USL - LSL)/(6*((mu - target)^2 + sigma^2)**0.5)
            
        if isLSL:
            Z_LSL = (mu - LSL)/sigma
            pNC_below = 1 - normsdist(Z_LSL)
            PPM_below = pNC_below*10**6
            metrics['p(N/C)-below'] = pNC_below
            metrics['Z-LSL'] = Z_LSL
            metrics['PPM-below'] = PPM_below
        if isUSL:
            Z_USL = (USL - mu)/sigma
            pNC_above = 1 - normsdist(Z_USL)
            PPM_above = pNC_above*10**6
            metrics['p(N/C)-above'] = pNC_above
            metrics['Z-USL'] = Z_USL
            metrics['PPM-above'] = PPM_above
        
        if isLSL and isUSL:
            pNC_total = pNC_below + pNC_above
            PPM_total = PPM_below + PPM_above
            metrics['p(N/C)-total'] = pNC_total
            metrics['PPM-total'] = PPM_total
            metrics['Zst'] = -normsinv(pNC_total)
            metrics['Zlt'] = -normsinv(pNC_total)

        return metrics
