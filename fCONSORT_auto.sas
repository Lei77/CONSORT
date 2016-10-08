
/****************************************************************************************************
* Program Name: fCONSORT_auto.sas 	 
* Purpose: 		Create macro to produce consort diagram by using provided template
* Written by: 	Lei Wang, Axio Research
* Date: 		07/14/2016	Created
				07/25/2016  Modify Reasons (frequency=0==>"Reasons NA")
				07/28/2016  Reasons: upper case ==> lower case
				08/25/2016  Add CutOffDate and RunDate to automatically update these two dates.
							(Use CUTOFFDATE and DATETIME in the template respectively)
*	
*****************************************************************************************************
*Key parameters:
   temp      - specify the template file location and name. 
			   Needs to be rtf file and should use AXIO RAS template setting for accesbling purpose
   outfig    - specify the output file location and name. Default location is &PROJROOT.\TLFs\.
   dataset   - the source dataset. i.e. analysis.adsl
   id        - record ID to be counted. Default is USUBJID;
   replaceX  - text in the template to be replaced by real data counts. X=1,2,3,4,5,6...
   multiX    - F for numbers, T for reasons. X=1,2,3,4,5,6...
   reasonX   - Variable that specifies the discontinue reasons. 
				if multiX=T, specify the reason variable (must be character). 
				if multiX=F, leave blank. X=1,2,3,4,5,6...
   criteriaX - specify the criteria (will be used under where) for counting each number. 
				i.e. randassi=1 and withdrew=0 . X=1,2,3,4,5,6...
   nreplace  - max X of replaceX 
****************************************************************************************************/



options NOQUOTELENMAX ;
%MACRO mfCONSORT (	temp =, 	
					outfig=&PROJROOT.\TLFs\,  
					dataset=Analysis.adsl,  
					id=USUBJID, 
					replace1=,  multi1=,  reason1=,  criteria1=, 
					replace2=,  multi2=,  reason2=,  criteria2=, 
					replace3=,  multi3=,  reason3=,  criteria3=, 
					replace4=,  multi4=,  reason4=,  criteria4=, 
					replace5=,  multi5=,  reason5=,  criteria5=, 
					replace6=,  multi6=,  reason6=,  criteria6=, 
					replace7=,  multi7=,  reason7=,  criteria7=, 
					replace8=,  multi8=,  reason8=,  criteria8=, 
					replace9=,  multi9=,  reason9=,  criteria9=, 
					replace10=, multi10=, reason10=, criteria10=, 
					replace11=, multi11=, reason11=, criteria11=, 
					replace12=, multi12=, reason12=, criteria12=, 
					replace13=, multi13=, reason13=, criteria13=, 
					replace14=, multi14=, reason14=, criteria14=, 
					replace15=, multi15=, reason15=, criteria15=, 
					replace16=, multi16=, reason16=, criteria16=, 
					replace17=, multi17=, reason17=, criteria17=, 
					replace18=, multi18=, reason18=, criteria18=, 
					replace19=, multi19=, reason19=, criteria19=, 
					replace20=, multi20=, reason20=, criteria20=, 
					replace21=, multi21=, reason21=, criteria21=, 
					replace22=, multi22=, reason22=, criteria22=, 
					replace23=, multi23=, reason23=, criteria23=, 
					replace24=, multi24=, reason24=, criteria24=, 
					replace25=, multi25=, reason25=, criteria25=, 
					replace26=, multi26=, reason26=, criteria26=, 
					replace27=, multi27=, reason27=, criteria27=, 
					replace28=, multi28=, reason28=, criteria28=, 
					replace29=, multi29=, reason29=, criteria29=, 
					replace30=, multi30=, reason30=, criteria30=, 
					replace31=, multi31=, reason31=, criteria31=, 
					replace32=, multi32=, reason32=, criteria32=, 
					replace33=, multi33=, reason33=, criteria33=, 
					replace34=, multi34=, reason34=, criteria34=, 
					replace35=, multi35=, reason35=, criteria35=, 
					replace36=, multi36=, reason36=, criteria36=, 
					replace37=, multi37=, reason37=, criteria37=, 
					replace38=, multi38=, reason38=, criteria38=, 
					replace39=, multi39=, reason39=, criteria39=, 
					replace40=, multi40=, reason40=, criteria40=, 
					replace41=, multi41=, reason41=, criteria41=, 
					replace42=, multi42=, reason42=, criteria42=, 
					replace43=, multi43=, reason43=, criteria43=, 
					replace44=, multi44=, reason44=, criteria44=, 
					replace45=, multi45=, reason45=, criteria45=, 
					replace46=, multi46=, reason46=, criteria46=, 
					replace47=, multi47=, reason47=, criteria47=, 
					replace48=, multi48=, reason48=, criteria48=, 
					replace49=, multi49=, reason49=, criteria49=, 
					replace50=, multi50=, reason50=, criteria50=, 

					nreplace= );

%LOCAL _rowreason 
	_number1  _number2  _number3  _number4  _number5  _number6  _number7  _number8  _number9  _number10 
	_number11 _number12 _number13 _number14 _number15 _number16 _number17 _number18 _number19 _number20
	_number21 _number22 _number23 _number24 _number25 _number26 _number27 _number28 _number29 _number30
	_number31 _number32 _number33 _number34 _number35 _number36 _number37 _number38 _number39 _number40
	_number41 _number42 _number43 _number44 _number45 _number46 _number47 _number48 _number49 _number50
	_checknumber1  _checknumber2  _checknumber3  _checknumber4  _checknumber5  _checknumber6  _checknumber7  _checknumber8  _checknumber9  _checknumber10 
	_checknumber11 _checknumber12 _checknumber13 _checknumber14 _checknumber15 _checknumber16 _checknumber17 _checknumber18 _checknumber19 _checknumber20
	_checknumber21 _checknumber22 _checknumber23 _checknumber24 _checknumber25 _checknumber26 _checknumber27 _checknumber28 _checknumber29 _checknumber30
	_checknumber31 _checknumber32 _checknumber33 _checknumber34 _checknumber35 _checknumber36 _checknumber37 _checknumber38 _checknumber39 _checknumber40
	_checknumber41 _checknumber42 _checknumber43 _checknumber44 _checknumber45 _checknumber46 _checknumber47 _checknumber48 _checknumber49 _checknumber50
;
 
filename confile1 "&temp "; 	* Import RTF template;
filename confile2 "&outfig"; 	* Export CONSORT diagram to a new RTF;

/*-----------------------------*/
/* Count numbers and reasons   */
/*-----------------------------*/ 

%DO i=1 %TO &nreplace;			
	%IF &&multi&i =F %THEN %DO;   	*---- Count numbers one by one ----;
		proc sql noprint;
			select count (distinct  &id) into: _number&i
			from &dataset
			where &&criteria&i;
		quit;
	%END;

	%IF &&multi&i=T %THEN %DO;		*---- Discontinue reasons ----;
		proc sql noprint;
			select count (distinct  &id) into: _checknumber&i
			from &dataset
			where &&criteria&i;
		quit;

		%IF &&_checknumber&i=0 %THEN %DO;  	* In case of frequency=0;
			%LET _number&i=• Reasons NA; 
		%END;

		%ELSE %DO;
		data _data1; set &dataset;
			&&reason&i=upcase(first(&&reason&i))||substr(lowcase(&&reason&i),2);
		run;

		proc freq data=_data1 order=freq; * Save the frequency table;
			table &&reason&i/noprint out=FreqCount ; 
			where &&criteria&i;
		run;

		data freqreason; 
			set FreqCount;  		* Concatenating reason and count;
			length _rowreason $1024;
			_rowreason = cats ("", "• ",&&reason&i) ||" " ||cats(" (N=",strip(put(count,8.)),")");
			where &&reason&i;
			keep _rowreason;
		run;

		proc transpose data=freqreason	
			out= trans_freq ; 		* Transpose;
			var _all_;		
		run;
		data transreason; 
			set trans_freq; 
			drop _name_; 
		run;

		proc sql noprint;			* Extract all the column names (all the reasons);
			select name into : _varlist&i separated by ','
			from dictionary.columns
			where libname = "WORK" and
			memname = "%upcase(transreason)";

		proc sql noprint;			* Connect with RTF unicode;
			select catx(" \par ", &&_varlist&i) length=1024  into: _number&i   
			from transreason;
		quit;

		%END;
	%END;
%END;

/*--------------------------------------------*/
/* Replacing numbers/reasons in RTF template  */
/*--------------------------------------------*/
data _null_;						
	infile confile1 lrecl=15000;
	input;
		%DO i=1 %TO &nreplace;
			%IF &&multi&i=F %THEN %DO;		* Number;
				_infile_ = transtrn(_infile_, "&&replace&i", trim(left(put(&&_number&i, 6.))));
			%END;

			%IF &&multi&i=T %THEN %DO;		* Reasons;
 				_infile_ = transtrn(_infile_, "&&replace&i", "&&_number&i");

			%END;
		%END;
		_infile_ = transtrn(_infile_, "CUTOFFDATE", "&cutoffdt"); *Use CUTOFFDATE in the headnote;
		_infile_ = transtrn(_infile_, "DATETIME", "&SYSDATE9 %sysfunc(time(),time5.0)"); *Use DATETIME in the footnote;
	file confile2 lrecl=15000;
	put _infile_;
run;

proc datasets library=work nolist; 			* Delete generated datasets ;
	delete FreqCount freqreason trans_freq transreason _data1;
run;

%mend mfCONSORT;







/*

*Example data;
* %include ;
%mfCONSORT (temp = fCONSORT_temp.rtf , 
			outfig=fCONSORT.rtf,
			pathdir=Q:\Shared\Lei Wang\CONSORT Diagram\,
			dataset=work.mydata , id=PatientKey , 
			replace1=XXX, 	multi1=F, 	reason1=, criteria1=randassi is not missing ,  
			replace2=LXX, 	multi2=F, 	reason2=, criteria2=randassi=1, 
			replace3=AXX, 	multi3=F, 	reason3=, criteria3=randassi=0,  
			replace4=LLX, 	multi4=F, 	reason4=, criteria4=randassi=1 and withdrew=1, 
			replace5=ALX, 	multi5=F, 	reason5=, criteria5=randassi=0 and withdrew=1,  
			replace6=RSNLFT, multi6=T, reason6=WhyWithdrew_chr, criteria6=randassi=1 and withdrew=1, 
			replace7=RSNRGT, multi7=T, reason7=WhyWithdrew_chr, criteria7=randassi=0 and withdrew=1, 
			replace8=TLRF,  multi8=F,  reason8=,  criteria8=randassi is NULL and flush3=1,  
			replace9=TLRP,  multi9=F,  reason9=,  criteria9=randassi is NULL and itch=1,  
			replace10=TLRO, multi10=F, reason10=, criteria10=randassi is NULL and arrhythmia=1 or glycControl=1 or OthIntol=1 ,  
			nreplace=10);


*/
