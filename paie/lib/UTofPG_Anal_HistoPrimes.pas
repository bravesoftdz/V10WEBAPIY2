{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 22/08/2002
Modifié le ... : 22/08/2002
Description .. : Gestion du tobviewer specifique à la gestion des
Suite ........ : primes
Suite ........ :
Mots clefs ... : PAIE
*****************************************************************}
{
PT1  : 03/05/2006 PH V_65 prise en compte driver dbMSSQL2005
}
unit UTofPG_Anal_HistoPrimes;

interface
uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,
      HTB97,Hqry, Stat,
      HCtrls,HEnt1,EntPaie,HMsgBox,UTOF,UTOB,AGLInit;

Type
     TOF_PG_Anal_HistoPrimes = Class (TOF)
      private
       public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnUpDate ;override;
     END ;

implementation

uses P5def,P5Util;

Function TransfSub (DD : String ; Deb,Long : integer) : String ;
BEGIN
dd:=dd+','+inttostr(deb)+','+inttostr(long) ;
Case V_PGI.Driver of
   dbSQLANY   : ; //A FAIRE        majtable
   dbSQLBASE  : ; //A FAIRE
   dbINTRBASE : Result:='MID('+dd+')' ;
   dbMSSQL, dbMSSQL2005, dbSYBASE    : Result:='substring('+dd+')' ;  // PT1
   dbORACLE7,dbORACLE8   : Result:='Substr('+dd+')' ;
   dbPOL      : Result:='Substr('+dd+')' ;
   dbMSACCESS : Result:='mid('+dd+')' ;
   dbDB2      : ; //A FAIRE
   dbINFORMIX : ; //A FAIRE
   else result:=dd ;
   END ;
END ;

procedure TOF_PG_Anal_HistoPrimes.OnArgument(Arguments: String);
var Ch1,Ch2,Cs1,Cs2    : String;
    Num            : Integer;
    Pref           : String;
begin
inherited ;
Pref := 'PSA';
if Pref <> '' then
   begin
   Cs1 := Pref+'_CODESTAT';
   Cs2 := 'T'+Pref+'_CODESTAT';
   VisibiliteStat (GetControl (Cs1),GetControl (Cs2)) ;
   VisibiliteStat (GetControl (Cs1+'_'),GetControl (Cs2+'_')) ;
   VisibiliteStat (GetControl (Cs1+'__'),GetControl (Cs2+'__')) ;
   end;
For Num := 1 to 4 do
 begin
 if Num>4 then Break;
 Ch1 := Pref+'_TRAVAILN'+IntToStr(Num);
 Ch2 := 'T'+Pref+'_TRAVAILN'+IntToStr(Num);
 if (Ch1 <> '') AND (Ch2 <> '') then
    begin
    VisibiliteChampSalarie (IntToStr(Num),GetControl (Ch1),GetControl (Ch2));
    VisibiliteChampSalarie (IntToStr(Num),GetControl (Ch1+'_'),GetControl (Ch2+'_'));
    VisibiliteChampSalarie (IntToStr(Num),GetControl (Ch1+'__'),GetControl (Ch2+'__'));
    end;
 end;

For Num := 1 to 4 do
 begin
 if Num>4 then Break;
 Ch1 := Pref+'_LIBREPCMB'+IntToStr(Num);
 Ch2 := 'T'+Pref+'_LIBREPCMB'+IntToStr(Num);
 if (Ch1 <> '') AND (Ch2 <> '') then
    begin
    VisibiliteChampLibreSal (IntToStr(Num),GetControl (Ch1),GetControl (Ch2));
    VisibiliteChampLibreSal (IntToStr(Num),GetControl (Ch1+'_'),GetControl (Ch2+'_'));
    VisibiliteChampLibreSal (IntToStr(Num),GetControl (Ch1+'__'),GetControl (Ch2+'__'));
    end;
 end;
For Num := 1 to 4 do
 begin
 if Num>4 then Break;
 Ch1 := Pref+'_BOOLLIBRE'+IntToStr(Num);
 Ch2 := 'T'+Pref+'_BOOLLIBRE'+IntToStr(Num);
 if (Ch1 <> '') AND (Ch2 <> '') then
    begin
    VisibiliteChampBoollX (IntToStr(Num),GetControl (Ch1));
    VisibiliteChampBoollX (IntToStr(Num),GetControl (Ch1+'_'));
    VisibiliteChampBoollX (IntToStr(Num),GetControl (Ch1+'__'));
    end;
 end;

end;

procedure TOF_PG_Anal_HistoPrimes.OnUpDate;
var StWhere : String;
    Pages   : TPageControl;
    StSql   : String;
    S1,S2   : String;
begin

S1 := TransfSub ('PHB_RUBRIQUE',5,1 );
S2 := TransfSub ('PHB_RUBRIQUE',1,4 );
Pages:=TPageControl(GetControl('Pages'));
stWhere := RecupWhereCritere(Pages);
StSql := 'SELECT PHB_SALARIE,PSA_LIBELLE,PSA_PRENOM,PHB_DATEFIN,PHB_RUBRIQUE,PHB_LIBELLE,PHB_BASEREM,PHB_TAUXREM,PHB_COEFFREM,PHB_MTREM FROM HISTOBULLETIN'+
   ' LEFT JOIN REMUNERATION ON PRM_RUBRIQUE = '+S2+' LEFT JOIN SALARIES ON PSA_SALARIE=PHB_SALARIE '+
   ' left join deportsal on PSE_SALARIE=PHB_SALARIE';
StWhere := StWhere + ' AND (PHB_MTREM<> 0 OR '+S1+'=".") AND '; // Pour prendre en compte les lignes de commentaire notament pour les absences
if NOT ConsultP then StWhere := StWhere+' PSE_Responsvar="'+LeSalarie+'" '
  else StWhere := StWhere+' PSE_ResponsABS="'+LeSalarie+'" ';
StSql := StSql +' '+StWhere+' ORDER BY PHB_SALARIE,PHB_RUBRIQUE,PHB_DATEFIN';
TFStat(Ecran).StSql := StSQL;
  inherited;
end;



Initialization
registerclasses([TOF_PG_Anal_HistoPrimes]);
end.


