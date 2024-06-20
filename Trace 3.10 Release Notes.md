## New Features
### Basics
- New 'calculation blocks' for buildings, pre-built for use in any calc sheet
	- *Room to room*
	- *Inside to outside* (including calculation as a plane source, and diffusivity correction as per *ISO 12354-4*)
- *Frequency Ranges* now includes *100Hz to 3150Hz*, (for Rw ratings and similar)
- *Frequency Band Cutoff* will now set the mode as an input
### Noise
- Added a *Fan Speed Correction* function and button (50*log(RPM1/RPM2))
- Added *Diesel & Gas-driven engines* to Estimation Methods
- Added *dB2Pa* for converting Decibels to Pressure
- Added *ISO23591_RTtargets* function, to return constants for the target range lines from the standard
- *Dir/Rev Sum* will now label the final total based on the line above your calculation block 
### Curves
- Separate *Inline rating* and ** button for:
	- *Rw+Ctr* rating
	- *STC* rating
	- *Lnw* rating
	- *AlphaW* rating
- Reference curves for:
	- *PNC*
	- *RC* (Mark II)
### Row functions
- *AutoSum* will now show the range to be summed, with user input to adjust as needed
- *Row reference* form will now reference **from** multiple selected rows **to** a destination row (but still works the other way if you want)
### Vibration
- *VC_rate* button now works for dB mode as well as mm/s mode
- Added a *DIN4150_Curve* function and a button 
- Added a *BS7385_Curve* function and a button 
- *AS2670* curves - dB mode updated with correct conversion factors
- Added some vibration prediction functions (but no buttons yet)
	- *Rayleigh Wave Loss (VibPrediction_RayleighWave)* 
	- *BS5228 (VibPrediction_BS5228)* 
	- *Amick method (VibPrediction_Amick)* 
## General improvements
- Created a *Loading Form* for use in various parts of the code. Used it in various parts of the code.
- *Extend* button is now more flexible, extending from the leftmost cell over as right as required (including dealing with excluded columns)
- *Direct-Reverberant Sum* now adds the result marker →
- *Schedule Builder* can summarise all the paths, using the result marker →
- *Schedule Builder* allows custom heading
- *Standard calc* form now resizes controls for long file names
- *Minus* marker added and used for *SPLMINUS*
- Markers put into groups
- Functions put into groups
- Cleaned up keytips (hint: press the ALT key!)
## Bugfixes
- *Ctr* function deals with negative
- 1-3 convert check sheet types -> CVT sheet was one band out (!)
- Fixed units on *Fan Simple*
- *ISO9613_Abar* method was applying the *Agr* factor in the wrong direction, due to differences in the sign convention
- *Import Fantech* now catches error for one-third octave sheets
- *LwFanSimple* now catches frequency bands outside the specified ranges
- Descriptions for *Parallelipiped* and *Conformal Surface* are now correct based on the (shared) form, not the ribbon click, as the user may have changed it.
- *VC Rating* will give error immediately if it's the wrong sheet layout