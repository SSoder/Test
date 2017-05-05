"""
User Defined Functions for the Valve Sizing Spreadsheet.
Functions written by Stefan Soder
Last Updated: 2016-07-29
"""
from xlpython import *
from iapws import IAPWS97

@xlfunc     #Written by Stefan Soder,2016-07-25 Last Update: 2016-07-29
def SteamType(pressure, temperature):
    
    #Convert psi to MPa; IAPWS97 object requires SI units    
    pressure = (pressure +14.7) * 0.006895
    #Convert Fahrenheit to Kelvin; IAPWS97 object requires SI units
    temperature = (((temperature-32)/1.8)+273)
    #Round the temperature to the nearest whole number.
    temperature = int(round(temperature))
    #Establishes an empty string to fill with the return value.
    fluid = ""
    #Instantiation of the steam object; assumes saturated steam.
    steam = IAPWS97(P=pressure, x=1)
    #Determines the saturation temperature of the steam.
    satsteamtemp = int(round(steam.T))
    """
    The logic below is to define the type of steam being used in the 
    customer's service. It assumes that the steam is saturated, and determines
    by the first if clause whether or not that is the case.
    It does so by comparing a ratio of the given line temperature to the
    saturation temperature determined above. If that ratio=1, then it is exactly
    the saturation temperature. The logic allows for a ratio right around one
    because customers are rarely going to give the perfect temperature for
    saturated steam at their pressure.
    """
    if 0.999 < temperature / satsteamtemp and temperature / satsteamtemp < 1.01:
            fluid = "Saturated Steam"
    elif 1.01 < temperature / satsteamtemp :
            fluid = "Superheated Steam"
    elif pressure > 22:
            fluid = "Supercritical Steam"
    else:   
        fluid = "Water"
    
    return (fluid)

@xlfunc     #Written by Stefan Soder,2016-07-25 Last Update: 2016-07-29
def SteamVolume(steamtype,pressure,temperature):
    #Initialize the volume variable.
    volume = 0
    #Convert pressure as in the SteamType function.
    pressure = (pressure + 14.7) * 0.006895
    #Convert Temperature as in the SteamType function
    temperature = (((temperature-32)/1.8)+273)
    temperature = int(round(temperature))
    """
    The logic below returns the specific volume attribute of the steam objects
    instatiated, while simultaneously converting them from MPa to psi.
    This is so that the returned value can be utilized correctly in the Valve
    Sizing Excel file.
    """
    if steamtype == "Saturated Steam":
        volsatsteam = IAPWS97(P=pressure, x=1)
        volume = volsatsteam.v * 16.0185
    elif steamtype == "Superheated Steam":
        volshsteam = IAPWS97(P=pressure, T=temperature)
        volume = volshsteam.v * 16.0185
    elif steamtype == "Supercritical Steam":
        volscsteam = IAPWS97(P=pressure, T=temperature)
        volume = volscsteam.v * 16.0185
    elif steamtype == "Water":
        volwater = IAPWS97(T=temperature,x=0)
        volume = volwater.v * 16.0185
    return(volume)

@xlfunc     #Written by Stefan Soder, 2016-07-27
def GasVolume(molecularweight, pressure, temperature):
    #Calculates the specific volume of a gas of given molecular weight. Uses
    #the ideal gas law to do so.
    absoluteRI = 1545.349 # (ft-lbs)/(lbmol*DegR)
    absoluteT = 459.67  # Conversion for Fahrenheit to Rankine
    absoluteP = 14.7    # Conversion to Absolute Pressure from Gage
    #Convert Temperature to Rankin
    temperature = round((absoluteT + temperature),3) 
    pressure = absoluteP + pressure #Determine Absolute Pressure from Gage
    #Calculating the Gas Constant for the given weight
    relativeR = round((absoluteRI / molecularweight),3)
    #And finally, calculate the specific volume of the gas in question.
    specificvolume = round(((temperature*relativeR)/(144*pressure)),6)
    return(specificvolume)
    
@xlfunc     #Written by Stefan Soder, 2016-07-29
def scfdConversion (scfd, pressure, temperature):
    #converts SCFD to gpm
    standardtemp = 519.67
    standardpres = 14.73
    absolutetemp = temperature + 459.67
    absolutepres = pressure + 14.73
    scfm = (scfd/24)/60
    pressure = absolutepres/standardpres
    temperature = standardtemp/absolutetemp
    acfm = scfm /(pressure*temperature)
    gpm = acfm * 7.481
    return(gpm)
    
    
            
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    