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
        
class UncertainVariable:
    """
    UncertainVariable objects track the effects of uncertainty, characterized 
    in terms of the first four standard moments of statistical distributions 
    (mean, variance, skewness and kurtosis coefficients). Monte Carlo simulation,
    in conjunction with Latin-hypercube based sampling performs the calculations.

    Parameters
    ----------
    rv : scipy.stats.distribution
        A distribution to characterize the uncertainty
    
    tag : str, optional
        A string identifier when information about this variable is printed to
        the screen
        
    Notes
    -----
    
    The ``scipy.stats`` module contains many distributions which we can use to
    perform any necessary uncertainty calculation. It is important to follow
    the initialization syntax for creating any kind of distribution object:
        
        - *Location* and *Scale* values must use the kwargs ``loc`` and 
          ``scale``
        - *Shape* values are passed in as non-keyword arguments before the 
          location and scale, (see below for syntax examples)..
        
    The mathematical operations that can be performed on uncertain objects will 
    work for any distribution supplied, but may be misleading if the supplied 
    moments or distribution is not accurately defined. Here are some guidelines 
    for creating UncertainVariable objects using some of the most common 
    statistical distributions:
    
    +---------------------------+-------------+-------------------+-----+---------+
    | Distribution              | scipy.stats |  args             | loc | scale   |
    |                           | class name  | (shape params)    |     |         |
    +===========================+=============+===================+=====+=========+
    | Normal(mu, sigma)         | norm        |                   | mu  | sigma   | 
    +---------------------------+-------------+-------------------+-----+---------+
    | Uniform(a, b)             | uniform     |                   | a   | b-a     |
    +---------------------------+-------------+-------------------+-----+---------+
    | Exponential(lamda)        | expon       |                   |     | 1/lamda |
    +---------------------------+-------------+-------------------+-----+---------+
    | Gamma(k, theta)           | gamma       | k                 |     | theta   |
    +---------------------------+-------------+-------------------+-----+---------+
    | Beta(alpha, beta, [a, b]) | beta        | alpha, beta       | a   | b-a     |
    +---------------------------+-------------+-------------------+-----+---------+
    | Log-Normal(mu, sigma)     | lognorm     | sigma             | mu  |         |
    +---------------------------+-------------+-------------------+-----+---------+
    | Chi-Square(k)             | chi2        | k                 |     |         |
    +---------------------------+-------------+-------------------+-----+---------+
    | F(d1, d2)                 | f           | d1, d2            |     |         |
    +---------------------------+-------------+-------------------+-----+---------+
    | Triangular(a, b, c)       | triang      | c                 | a   | b-a     |
    +---------------------------+-------------+-------------------+-----+---------+
    | Student-T(v)              | t           | v                 |     |         |
    +---------------------------+-------------+-------------------+-----+---------+
    | Weibull(lamda, k)         | exponweib   | lamda, k          |     |         |
    +---------------------------+-------------+-------------------+-----+---------+
    | Bernoulli(p)              | bernoulli   | p                 |     |         |
    +---------------------------+-------------+-------------------+-----+---------+
    | Binomial(n, p)            | binomial    | n, p              |     |         |
    +---------------------------+-------------+-------------------+-----+---------+
    | Geometric(p)              | geom        | p                 |     |         |
    +---------------------------+-------------+-------------------+-----+---------+
    | Hypergeometric(N, n, K)   | hypergeom   | N, n, K           |     |         |
    +---------------------------+-------------+-------------------+-----+---------+
    | Poisson(lamda)            | poisson     | lamda             |     |         |
    +---------------------------+-------------+-------------------+-----+---------+
    
    Thus, each distribution above would have the same call signature::
        
        >>> import scipy.stats as ss
        >>> ss.your_dist_here(args, loc=loc, scale=scale)
        
    ANY SCIPY.STATS.DISTRIBUTION SHOULD WORK! IF ONE DOESN'T, PLEASE LET ME
    KNOW!
    
    Convenient constructors have been created to make assigning these 
    distributions easier. They follow the parameter notation found in the
    respective Wikipedia articles:
    
    +---------------------------+---------------------------------------------------------------+
    | MCERP Distibution         | Wikipedia page                                                |
    +===========================+===============================================================+
    | N(mu, sigma)              | http://en.wikipedia.org/wiki/Normal_distribution              |
    +---------------------------+---------------------------------------------------------------+
    | U(a, b)                   | http://en.wikipedia.org/wiki/Uniform_distribution_(continuous)|
    +---------------------------+---------------------------------------------------------------+
    | Exp(lamda, [mu])          | http://en.wikipedia.org/wiki/Exponential_distribution         |
    +---------------------------+---------------------------------------------------------------+
    | Gamma(k, theta)           | http://en.wikipedia.org/wiki/Gamma_distribution               |
    +---------------------------+---------------------------------------------------------------+
    | Beta(alpha, beta, [a, b]) | http://en.wikipedia.org/wiki/Beta_distribution                |
    +---------------------------+---------------------------------------------------------------+
    | LogN(mu, sigma)           | http://en.wikipedia.org/wiki/Log-normal_distribution          |
    +---------------------------+---------------------------------------------------------------+
    | X2(df)                    | http://en.wikipedia.org/wiki/Chi-squared_distribution         |
    +---------------------------+---------------------------------------------------------------+
    | F(dfn, dfd)               | http://en.wikipedia.org/wiki/F-distribution                   |
    +---------------------------+---------------------------------------------------------------+
    | Tri(a, b, c)              | http://en.wikipedia.org/wiki/Triangular_distribution          |
    +---------------------------+---------------------------------------------------------------+
    | T(df)                     | http://en.wikipedia.org/wiki/Student's_t-distribution         |
    +---------------------------+---------------------------------------------------------------+
    | Weib(lamda, k)            | http://en.wikipedia.org/wiki/Weibull_distribution             |
    +---------------------------+---------------------------------------------------------------+
    | Bern(p)                   | http://en.wikipedia.org/wiki/Bernoulli_distribution           |
    +---------------------------+---------------------------------------------------------------+
    | B(n, p)                   | http://en.wikipedia.org/wiki/Binomial_distribution            |
    +---------------------------+---------------------------------------------------------------+
    | G(p)                      | http://en.wikipedia.org/wiki/Geometric_distribution           |
    +---------------------------+---------------------------------------------------------------+
    | H(M, n, N)                | http://en.wikipedia.org/wiki/Hypergeometric_distribution      |
    +---------------------------+---------------------------------------------------------------+
    | Pois(lamda)               | http://en.wikipedia.org/wiki/Poisson_distribution             |
    +---------------------------+---------------------------------------------------------------+

    Thus, the following are equivalent::

        >>> x = N(10, 1)
        >>> x = uv(ss.norm(loc=10, scale=1))

    Examples
    --------
    A three-part assembly
        
        >>> x1 = N(24, 1)
        >>> x2 = N(37, 4)
        >>> x3 = Exp(2)  # Exp(mu=0.5) works too
        
        >>> Z = (x1*x2**2)/(15*(1.5 + x3))
        >>> Z
        uv(1161.46231679, 116646.762981, 0.345533974771, 3.00791101068)

    The result shows the mean, variance, and standardized skewness and kurtosis
    of the output variable Z, which will vary from use to use due to the random
    nature of Monte Carlo simulation and latin-hypercube sampling techniques.
    
    Basic math operations may be applied to distributions, where all 
    statistical calculations are performed using latin-hypercube enhanced Monte
    Carlo simulation. Nearly all of the built-in trigonometric-, logarithm-, 
    etc. functions of the ``math`` module have uncertainty-compatible 
    counterparts that should be used when possible since they support both 
    scalar values and uncertain objects. These can be used after importing the 
    ``umath`` module::
        
        >>> from mcerp.umath import * # sin(), sqrt(), etc.
        >>> sqrt(x1)
        uv(4.89791765647, 0.0104291897681, -0.0614940614672, 3.00264937735)
    
    At any time, the standardized statistics can be retrieved using::
        
        >>> x1.mean
        >>> x1.var  # x1.std (standard deviation) is also available
        >>> x1.skew
        >>> x1.kurt
    
    or all four together with::
    
        >>> x1.stats
    
    By default, the Monte Carlo simulation uses 10000 samples, but this can be
    changed at any time with::
        
        >>> mcerp.npts = number_of_samples
    
    Any value from 1,000 to 1,000,000 is recommended (more samples means more
    accurate, but also means more time required to perform the calculations). 
    Although it can be changed, since variables retain their samples from one
    calculation to the next, this parameter should be changed before any 
    calculations are performed to ensure parameter compatibility (this may 
    change to be more dynamic in the future, but for now this is how it is).
    
    Also, to see the underlying distribution of the variable, and if matplotlib
    is installed, simply call its plot method::
        
        >>> x1.plot()
    
    Optional kwargs can be any valid kwarg used by matplotlib.pyplot.plot
    
    See Also
    --------
    N, U, Exp, Gamma, Beta, LogN, X2, F, Tri, T, Weib, Bern, B, G, H, Pois
        
    """
    
    def __init__(self, rv, tag=None):
        
        assert hasattr(rv, 'dist'), 'Input must be a  distribution from ' + \
            'the scipy.stats module.'
        self.rv = rv
        
        # generate the latin-hypercube points
        self._mcpts = lhd(dist=self.rv, size=npts).flatten()
        self.tag = tag
        
    def plot(self, hist=False, **kwargs):
        """
        Plot the distribution of the UncertainVariable. Continuous 
        distributions are plotted with a line plot and discrete distributions
        are plotted with discrete circles.
        
        Optional
        --------
        hist : bool
            If true, a histogram is displayed
        kwargs : any valid matplotlib.pyplot.plot kwarg
        
        """
        
        if hist:
            vals = self._mcpts
            low = vals.min()
            high = vals.max()
            h = plt.hist(vals, bins=np.round(np.sqrt(len(vals))), 
                     histtype='stepfilled', normed=True, **kwargs)

            if self.tag is not None:
                # plt.suptitle('Histogram of (' + self.tag + ')')
                plt.title(str(self), fontsize=12)
            else:
                # plt.suptitle('Histogram of')
                plt.title(str(self), fontsize=12)

            plt.ylim(0, 1.1*h[0].max())
        else:
            bound = 0.0001
            low = self.rv.ppf(bound)
            high = self.rv.ppf(1 - bound)
            if hasattr(self.rv.dist, 'pmf'):
                low = int(low)
                high = int(high)
                vals = range(low, high + 1)
                plt.plot(vals, self.rv.pmf(vals), 'o', **kwargs)

                if self.tag is not None:
                    # plt.suptitle('PMF of (' + self.tag + ')')
                    plt.title(str(self), fontsize=12)
                else:
                    # plt.suptitle('PMF of')
                    plt.title(str(self), fontsize=12)

            else:
                vals = np.linspace(low, high, 500)
                plt.plot(vals, self.rv.pdf(vals), **kwargs)

                if self.tag is not None:
                    # plt.suptitle('PDF of ('+self.tag+')')
                    plt.title(str(self), fontsize=12)
                else:
                    # plt.suptitle('PDF of')
                    plt.title(str(self), fontsize=12)

        plt.xlim(low - (high - low)*0.1, high + (high - low)*0.1)

                
        
uv = UncertainVariable # a nicer form for the user

# DON'T MOVE THIS IMPORT!!! The prior definitions must be in place before
# importing the correlation-related functions
from correlate import *

###############################################################################
# Define some convenience constructors for common statistical distributions.
# Hopefully these are a little easier/more intuitive to use than the 
# scipy.stats.distributions.
###############################################################################

def N(mu, sigma, tag=None):
    """
    A Normal (or Gaussian) random variate
    
    Parameters
    ----------
    mu : scalar
        The mean value of the distribution
    sigma : scalar
        The standard deviation (must be positive and non-zero)
    """
    assert sigma>0, 'Normal "sigma" must be greater than zero'
    return uv(ss.norm(loc=mu, scale=sigma), tag=tag)

###############################################################################

def U(a, b, tag=None):
    """
    A Uniform random variate
    
    Parameters
    ----------
    a : scalar
        Lower bound of the distribution support.
    b : scalar
        Upper bound of the distribution support.
    """
    assert a<b, 'Uniform "a" must be less than "b"'
    return uv(ss.uniform(loc=a, scale=b-a), tag=tag)

###############################################################################

def Exp(lamda, tag=None):
    """
    An Exponential random variate
    
    Parameters
    ----------
    lamda : scalar
        The inverse scale (as shown on Wikipedia). (FYI: mu = 1/lamda.)
    """
    assert lamda>0, 'Exponential "lamda" must be greater than zero'
    return uv(ss.expon(scale=1./lamda), tag=tag)

###############################################################################

def Gamma(k, theta, tag=None):
    """
    A Gamma random variate
    
    Parameters
    ----------
    k : scalar
        The shape parameter (must be positive and non-zero)
    theta : scalar
        The scale parameter (must be positive and non-zero)
    """
    assert k>0 and theta>0, 'Gamma "k" and "theta" parameters must be greater than zero'
    return uv(ss.gamma(k, scale=theta), tag=tag)

###############################################################################

def Beta(alpha, beta, a=0, b=1, tag=None):
    """
    A Beta random variate
    
    Parameters
    ----------
    alpha : scalar
        The first shape parameter
    beta : scalar
        The second shape parameter
    
    Optional
    --------
    a : scalar
        Lower bound of the distribution support (default=0)
    b : scalar
        Upper bound of the distribution support (default=1)
    """
    assert alpha>0 and beta>0, 'Beta "alpha" and "beta" parameters must be greater than zero'
    return uv(ss.beta(alpha, beta, loc=a, scale=b-a), tag=tag)

###############################################################################

def LogN(mu, sigma, tag=None):
    """
    A Log-Normal random variate
    
    Parameters
    ----------
    mu : scalar
        The location parameter
    sigma : scalar
        The scale parameter (must be positive and non-zero)
    """
    assert sigma>0, 'Log-Normal "sigma" must be positive'
    return uv(ss.lognorm(sigma, loc=mu), tag=tag)

###############################################################################

def Chi2(k, tag=None):
    """
    A Chi-Squared random variate
    
    Parameters
    ----------
    k : int
        The degrees of freedom of the distribution (must be greater than one)
    """
    assert isinstance(k, int) and k>=1, 'Chi-Squared "k" must be an integer greater than 0'
    return uv(ss.chi2(k), tag=tag)

###############################################################################

def F(d1, d2, tag=None):
    """
    An F (fisher) random variate
    
    Parameters
    ----------
    d1 : int
        Numerator degrees of freedom
    d2 : int
        Denominator degrees of freedom
    """
    assert isinstance(d1, int) and d1>=1, 'Fisher (F) "d1" must be an integer greater than 0'
    assert isinstance(d2, int) and d2>=1, 'Fisher (F) "d2" must be an integer greater than 0'
    return uv(ss.f(d1, d2), tag=tag)

###############################################################################

def Tri(a, b, c, tag=None):
    """
    A triangular random variate
    
    Parameters
    ----------
    a : scalar
        Lower bound of the distribution support
    b : scalar
        Upper bound of the distribution support
    c : scalar
        The location of the triangle's peak (a <= c <= b)
    """
    assert a<=c<=b, 'Triangular "c" must lie between "a" and "b"'
    return uv(ss.triang((1.0*c-a)/(b-a), loc=a, scale=b-a), tag=tag)

###############################################################################

def T(v, tag=None):
    """
    A Student-T random variate
    
    Parameters
    ----------
    v : int
        The degrees of freedom of the distribution (must be greater than one)
    """
    assert isinstance(v, int) and v>=1, 'Student-T "v" must be an integer greater than 0'
    return uv(ss.t(v), tag=tag)

###############################################################################

def Weib(lamda, k, tag=None):
    """
    A Weibull random variate
    
    Parameters
    ----------
    lamda : scalar
        The scale parameter
    k : scalar
        The shape parameter
    """
    assert lamda>0 and k>0, 'Weibull "lamda" and "k" parameters must be greater than zero'
    return uv(ss.exponweib(lamda, k), tag=tag)

###############################################################################

def Bern(p, tag=None):
    """
    A Bernoulli random variate
    
    Parameters
    ----------
    p : scalar
        The probability of success
    """
    assert 0<p<1, 'Bernoulli probability "p" must be between zero and one, non-inclusive'
    return uv(ss.bernoulli(p), tag=tag)

###############################################################################

def B(n, p, tag=None):
    """
    A Binomial random variate
    
    Parameters
    ----------
    n : int
        The number of trials
    p : scalar
        The probability of success
    """
    assert int(n)==n and n>0, 'Binomial number of trials "n" must be an integer greater than zero'
    assert 0<p<1, 'Binomial probability "p" must be between zero and one, non-inclusive'
    return uv(ss.binom(n, p), tag=tag)

###############################################################################

def G(p, tag=None):
    """
    A Geometric random variate
    
    Parameters
    ----------
    p : scalar
        The probability of success
    """
    assert 0<p<1, 'Geometric probability "p" must be between zero and one, non-inclusive'
    return uv(ss.geom(p), tag=tag)

###############################################################################

def H(N, n, K, tag=None):
    """
    A Hypergeometric random variate
    
    Parameters
    ----------
    N : int
        The total population size
    n : int
        The number of individuals of interest in the population
    K : int
        The number of individuals that will be chosen from the population
        
    Example
    -------
    (Taken from the wikipedia page) Assume we have an urn with two types of
    marbles, 45 black ones and 5 white ones. Standing next to the urn, you
    close your eyes and draw 10 marbles without replacement. What is the
    probability that exactly 4 of the 10 are white?
    ::
    
        >>> black = 45
        >>> white = 5
        >>> draw = 10
        
        # Now we create the distribution
        >>> h = H(black + white, white, draw)
        
        # To check the probability, in this case, we can use the underlying
        #  scipy.stats object
        >>> h.rv.pmf(4)  # What is the probability that white count = 4?
        0.0039645830580151975
        
    """
    assert int(N)==N and N>0, 'Hypergeometric total population size "N" must be an integer greater than zero.'
    assert int(n)==n and 0<n<=N, 'Hypergeometric interest population size "n" must be an integer greater than zero and no more than the total population size.'
    assert int(K)==K and 0<K<=N, 'Hypergeometric chosen population size "K" must be an integer greater than zero and no more than the total population size.'
    return uv(ss.hypergeom(N, n, K), tag=tag)

###############################################################################

def Pois(lamda, tag=None):
    """
    A Poisson random variate
    
    Parameters
    ----------
    lamda : scalar
        The rate of an occurance within a specified interval of time or space.
    """
    assert lamda>0, 'Poisson "lamda" must be greater than zero.'
    return uv(ss.poisson(lamda), tag=tag)

###############################################################################

def covariance_matrix(nums_with_uncert):
    """
    Calculate the covariance matrix of uncertain variables, oriented by the
    order of the inputs
    
    Parameters
    ----------
    nums_with_uncert : array-like
        A list of variables that have an associated uncertainty
    
    Returns
    -------
    cov_matrix : 2d-array-like
        A nested list containing covariance values
    
    Example
    -------
    
        >>> x = N(1, 0.1)
        >>> y = N(10, 0.1)
        >>> z = x + 2*y
        >>> covariance_matrix([x,y,z])
        [[  9.99694861e-03   2.54000840e-05   1.00477488e-02]
         [  2.54000840e-05   9.99823207e-03   2.00218642e-02]
         [  1.00477488e-02   2.00218642e-02   5.00914772e-02]]

    """
    ufuncs = map(to_uncertain_func,nums_with_uncert)
    cov_matrix = []
    for (i1, expr1) in enumerate(ufuncs):
        coefs_expr1 = []
        mean1 = expr1.mean
        for (i2, expr2) in enumerate(ufuncs[:i1+1]):
            mean2 = expr2.mean
            coef = np.mean((expr1._mcpts - mean1)*(expr2._mcpts - mean2))
            coefs_expr1.append(coef)
        cov_matrix.append(coefs_expr1)
        
    # We symmetrize the matrix:
    for (i, covariance_coefs) in enumerate(cov_matrix):
        covariance_coefs.extend(cov_matrix[j][i]
                                for j in range(i+1, len(cov_matrix)))

    return cov_matrix

def correlation_matrix(nums_with_uncert):
    """
    Calculate the correlation matrix of uncertain variables, oriented by the
    order of the inputs
    
    Parameters
    ----------
    nums_with_uncert : array-like
        A list of variables that have an associated uncertainty
    
    Returns
    -------
    corr_matrix : 2d-array-like
        A nested list containing covariance values
    
    Example
    -------
    
        >>> x = N(1, 0.1)
        >>> y = N(10, 0.1)
        >>> z = x + 2*y
        >>> correlation_matrix([x,y,z])
        [[ 0.99969486  0.00254001  0.4489385 ]
         [ 0.00254001  0.99982321  0.89458702]
         [ 0.4489385   0.89458702  1.        ]]

    """
    ufuncs = map(to_uncertain_func, nums_with_uncert)
    data = np.vstack([ufunc._mcpts for ufunc in ufuncs])
    return np.corrcoef(data.T, rowvar=0)    

