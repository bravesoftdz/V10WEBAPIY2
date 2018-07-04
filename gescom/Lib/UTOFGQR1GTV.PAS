{***********UNITE*************************************************
Auteur  ...... : LUQUET Jean-Marc
Créé le ...... : 05/09/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : GQR1GTV ()
Mots clefs ... : TOF;GQR1GTV
*****************************************************************}
Unit UTOFGQR1GTV ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils,  ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF,stat, ed_tools ,
     Dialogs ;

Type
  TOF_GQR1GTV = Class (TOF)
    procedure OnLoad                   ; override ;
  end ;
  TOF_GQR1GPA = Class (TOF)
    procedure OnLoad                   ; override ;
  end ;
  TOF_GQR1GII = Class (TOF)
   private
    ListeAgence,ListeMarche : string;
    Deb_Saisie, Fin_Saisie,Deb_Moins12,Fin_Moins12,Deb_Plus4,Fin_Plus4 : TDateTime;
    procedure RequeteParAgence;
    procedure RequeteParMarche;
    Function CritereAgence(Requete,champ :string):string ;
    Function CritereMarche(Requete,champ :string):string ;
   public
    procedure OnLoad                   ; override ;
  end ;

Implementation

procedure TOF_GQR1GTV.OnLoad ;
var Stdate : string;
    Requete,Requete0,Requete1,Requete2,Requete3,Requete4,Requete5,Requete6,Requete7 : string;
    Requete8,Requete9,Requete10,Requete11 : string;
    CRI_MARCHE,CRI_GPMARCHE,CRI_YTCMARCHE,CRI_AGENCE,CRI_GPAGENCE,CRI_YTCAGENCE : string;
    Deb_Saisie, Fin_Saisie,Deb_Moins12,Fin_Moins12,Deb_Plus4,Fin_Plus4,Deb_Mois,Fin_Mois : TDateTime;
    listeCode,ListeIn : string;

begin
  Inherited ;
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD") IS NOT NULL) DROP TABLE #PRD ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD1") IS NOT NULL) DROP TABLE #PRD1 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD2") IS NOT NULL) DROP TABLE #PRD2 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD3") IS NOT NULL) DROP TABLE #PRD3 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD4") IS NOT NULL) DROP TABLE #PRD4 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD5") IS NOT NULL) DROP TABLE #PRD5 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD6") IS NOT NULL) DROP TABLE #PRD6 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD7") IS NOT NULL) DROP TABLE #PRD7 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD8") IS NOT NULL) DROP TABLE #PRD8 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD9") IS NOT NULL) DROP TABLE #PRD9 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD10") IS NOT NULL) DROP TABLE #PRD10 ');

  Deb_Mois := DebutDeMois (V_PGI.DateEntree);
  Fin_Mois := FinDeMois (V_PGI.DateEntree);
  Stdate := GetControlText('GPDATEPIECE0');
  if (IsValidDate(StDate)) then Deb_Saisie := DebutDeMois(StrToDate (StDate))
  else  Deb_Saisie := Deb_Mois;

  Stdate := GetControlText('GPDATEPIECE0_');
  if (IsValidDate(StDate)) then Fin_Saisie := FinDeMois(StrToDate(StDate))
  else Fin_Saisie := Fin_Mois;
  if (Fin_Saisie < Deb_Saisie) then Fin_Saisie := FinDeMois (Deb_Saisie);

  if  (Fin_Saisie >= V_PGI.DateEntree) then
  begin
    Fin_Moins12 := FinDeMois(PlusMois (Fin_Saisie, -1));
    Deb_Moins12 := DebutDeMois(PlusMois (Fin_Moins12,-11));
    Deb_Plus4 := DebutDeMois ( Fin_Saisie);
    Fin_Plus4 := FinDeMois(PlusMois (Deb_Plus4,3));
  end else
  begin
    Fin_Moins12 := Fin_Saisie ;
    Deb_Moins12 := DebutDeMois(PlusMois (Fin_Moins12,-11));
    Deb_Plus4 := DebutDeMois ( PlusMois (Fin_Saisie,1));
    Fin_Plus4 := FinDeMois(PlusMois (Deb_Plus4,3));
  end;
  // MAJ requête de la fiche
  SetControlText( 'GPDATEPIECE0' , DateToStr (Deb_Saisie));
  SetControlText( 'GPDATEPIECE0_' , DateToStr (Fin_Saisie));

  ListeCode := THMultiValcomboBox(GetControl('CODEMARCHE')).value  ;
  if (trim(ListeCode)<>'') then
  begin
    listeIn := 'IN ("'+ReadtokenSt(ListeCode)+'"';
    while (trim(ListeCode)<>'') do listeIn := listeIn+',"'+ReadtokenSt(ListeCode)+'"';
    listeIn := listeIn + ')';
    CRI_MARCHE :=' CC1.YX_code ' + listeIn ;
    CRI_GPMARCHE := ' gp_libretiers1 ' + listeIn ;
    CRI_YTCMARCHE := ' ytc_tablelibretiers1 ' + listeIn ;
  end else
  begin
    CRI_MARCHE := '1=1';      
    CRI_GPMARCHE := '1=1';
    CRI_YTCMARCHE := '1=1';
  end;

  ListeCode := THMultiValcomboBox(GetControl('CODEAGENCE')).value  ;
  if (trim(ListeCode)<>'') then
  begin
  listeIn := 'IN ("'+ReadtokenSt(ListeCode)+'"';
  while (trim(ListeCode)<>'') do listeIn := listeIn+',"'+ReadtokenSt(ListeCode)+'"';
  listeIn := listeIn + ')';
  CRI_AGENCE :=' CC2.YX_code ' + listeIn ;
  CRI_GPAGENCE := ' gp_libretiers2 ' + listeIn ;
  CRI_YTCAGENCE := ' ytc_tablelibretiers2 ' + listeIn ;
  end else
  begin
  CRI_AGENCE := '1=1';
  CRI_GPAGENCE := '1=1';     
  CRI_YTCAGENCE := '1=1';
  end;

  Requete0 := getcontroltext('REQUETE0');
  Requete0 :=  StringReplace( Requete0,'[CRI_MARCHE]',CRI_MARCHE,[rfIgnoreCase]);
  Requete0 :=  StringReplace( Requete0,'[CRI_AGENCE]',CRI_AGENCE,[rfIgnoreCase]);
  ExecuteSQL(Requete0);
  //CRI_MARCHE,CRI_GPMARCHE,CRI_YTCMARCHE,CRI_AGENCE,CRI_GPAGENCE,CRI_YTCAGENCE
  Requete1 := getcontroltext('REQUETE1');
  Requete1 := StringReplace(Requete1, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete1 := StringReplace(Requete1, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete1 := StringReplace( Requete1,'[CRI_GPMARCHE]',CRI_GPMARCHE,[rfIgnoreCase]);
  Requete1 := StringReplace( Requete1,'[CRI_GPAGENCE]',CRI_GPAGENCE,[rfIgnoreCase]);
  ExecuteSQL(Requete1);
  Requete2 := getcontroltext('REQUETE2');
  Requete2 := StringReplace(Requete2, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete2 := StringReplace(Requete2, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete2 := StringReplace( Requete2,'[CRI_GPMARCHE]',CRI_GPMARCHE,[rfIgnoreCase]);
  Requete2 := StringReplace( Requete2,'[CRI_GPAGENCE]',CRI_GPAGENCE,[rfIgnoreCase]);
  ExecuteSQL(Requete2);
  Requete3 := getcontroltext('REQUETE3');
  Requete3 := StringReplace(Requete3, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete3 := StringReplace(Requete3, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete3 := StringReplace( Requete3,'[CRI_GPMARCHE]',CRI_GPMARCHE,[rfIgnoreCase]);
  Requete3 := StringReplace( Requete3,'[CRI_GPAGENCE]',CRI_GPAGENCE,[rfIgnoreCase]);
  ExecuteSQL(Requete3);
  Requete4 := getcontroltext('REQUETE4');
  Requete4 := StringReplace(Requete4, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete4 := StringReplace(Requete4, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete4 := StringReplace( Requete4,'[CRI_YTCMARCHE]',CRI_YTCMARCHE,[rfIgnoreCase]);
  Requete4 := StringReplace( Requete4,'[CRI_YTCAGENCE]',CRI_YTCAGENCE,[rfIgnoreCase]);
  ExecuteSQL(Requete4);
  Requete5 := getcontroltext('REQUETE5');
  Requete5 := StringReplace(Requete5, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete5 := StringReplace(Requete5, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete5 := StringReplace( Requete5,'[CRI_YTCMARCHE]',CRI_YTCMARCHE,[rfIgnoreCase]);
  Requete5 := StringReplace( Requete5,'[CRI_YTCAGENCE]',CRI_YTCAGENCE,[rfIgnoreCase]);
  ExecuteSQL(Requete5);
  Requete6 := getcontroltext('REQUETE6');
  Requete6 := StringReplace(Requete6, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete6 := StringReplace(Requete6, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete6 := StringReplace( Requete6,'[CRI_YTCMARCHE]',CRI_YTCMARCHE,[rfIgnoreCase]);
  Requete6 := StringReplace( Requete6,'[CRI_YTCAGENCE]',CRI_YTCAGENCE,[rfIgnoreCase]);
  ExecuteSQL(Requete6);
  Requete7 := getcontroltext('REQUETE7');
  Requete7 := StringReplace(Requete7, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete7 := StringReplace(Requete7, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete7 := StringReplace( Requete7,'[CRI_YTCMARCHE]',CRI_YTCMARCHE,[rfIgnoreCase]);
  Requete7 := StringReplace( Requete7,'[CRI_YTCAGENCE]',CRI_YTCAGENCE,[rfIgnoreCase]);
  ExecuteSQL(Requete7);
  Requete8 := getcontroltext('REQUETE8');
  Requete8 := StringReplace(Requete8, '[CRI_DEBPERSPOST4M]', USDateTime(Deb_Plus4), [rfReplaceAll]);
  Requete8 := StringReplace(Requete8, '[CRI_FINPERSPOST4M]', USDateTime(Fin_Plus4), [rfReplaceAll]);
  Requete8 := StringReplace( Requete8,'[CRI_YTCMARCHE]',CRI_YTCMARCHE,[rfIgnoreCase]);
  Requete8 := StringReplace( Requete8,'[CRI_YTCAGENCE]',CRI_YTCAGENCE,[rfIgnoreCase]);
  ExecuteSQL(Requete8);
  Requete9 := getcontroltext('REQUETE9');
  Requete9 := StringReplace(Requete9, '[CRI_DEBPERSANT12M]', USDateTime(Deb_Moins12), [rfReplaceAll]);
  Requete9 := StringReplace(Requete9, '[CRI_FINPERSANT12M]', USDateTime(Fin_Moins12), [rfReplaceAll]);
  Requete9 := StringReplace( Requete9,'[CRI_YTCMARCHE]',CRI_YTCMARCHE,[rfIgnoreCase]);
  Requete9 := StringReplace( Requete9,'[CRI_YTCAGENCE]',CRI_YTCAGENCE,[rfIgnoreCase]);
  ExecuteSQL(Requete9);
  Requete10 := getcontroltext('REQUETE10');
  Requete10 := StringReplace(Requete10, '[CRI_DEBMOISCOURS]', USDateTime(Deb_Mois), [rfReplaceAll]);
  Requete10 := StringReplace(Requete10, '[CRI_FINMOISCOURS]', USDateTime(Fin_Mois), [rfReplaceAll]);
  Requete10 := StringReplace( Requete10,'[CRI_YTCMARCHE]',CRI_YTCMARCHE,[rfIgnoreCase]);
  Requete10 := StringReplace( Requete10,'[CRI_YTCAGENCE]',CRI_YTCAGENCE,[rfIgnoreCase]);
  ExecuteSQL(Requete10);

  Requete := getcontroltext('REQUETE');
  Requete11 := getcontroltext('REQUETE11');
  TFStat(Ecran).FSQL.lines.clear;
  TFStat(Ecran).FSQL.lines[0] := Requete +#$D#$A;
  TFStat(Ecran).FSQL.lines[1] := Requete11 +#$D#$A;
end;
procedure TOF_GQR1GPA.OnLoad ;
var Stdate : string;
    Requete1,Requete2,Requete3,Requete4,Requete5 : string;
    Deb_Saisie, Fin_Saisie,Deb_Moins1,Fin_Moins1,Deb_Plus4,Fin_Plus4,Deb_Mois,Fin_Mois : TDateTime;
begin
  Inherited ;
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD1") IS NOT NULL) DROP TABLE #PRD1 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD2") IS NOT NULL) DROP TABLE #PRD2 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD3") IS NOT NULL) DROP TABLE #PRD3 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD4") IS NOT NULL) DROP TABLE #PRD4 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD5") IS NOT NULL) DROP TABLE #PRD5 ');

  Deb_Mois := DebutDeMois (V_PGI.DateEntree);
  Fin_Mois := FinDeMois (V_PGI.DateEntree);
  Stdate := GetControlText('DATEMOISPREV');
  if (IsValidDate(StDate)) then Deb_Saisie := DebutDeMois(StrToDate (StDate))
  else  Deb_Saisie := Deb_Mois;

  Stdate := GetControlText('DATEMOISPREV_');
  if (IsValidDate(StDate)) then Fin_Saisie := FinDeMois(StrToDate(StDate))
  else Fin_Saisie := Fin_Mois;
  if (Fin_Saisie < Deb_Saisie) then Fin_Saisie := FinDeMois (Deb_Saisie);

  Deb_Moins1 := DebutDeMois(PlusMois (Deb_Saisie,-1));
  Fin_Moins1 := FinDeMois(Deb_Moins1);

  if  (Fin_Saisie >= V_PGI.DateEntree) then
  begin
    Deb_Plus4 := DebutDeMois ( Fin_Saisie);
    Fin_Plus4 := FinDeMois(PlusMois (Deb_Plus4,3));
  end else
  begin
    Deb_Plus4 := DebutDeMois ( PlusMois (Fin_Saisie,1));
    Fin_Plus4 := FinDeMois(PlusMois (Deb_Plus4,3));
  end;
  // MAJ requête de la fiche
  SetControlText( 'DATEMOISPREV' , DateToStr (Deb_Saisie));
  SetControlText( 'DATEMOISPREV_' , DateToStr (Fin_Saisie));
  SetControlText( 'DATEMOISANT' , DateToStr (Deb_Moins1));
  SetControlText( 'DATEMOISANT_' , DateToStr (Fin_Moins1));
  SetControlText( 'DATEDEB4M' , DateToStr (Deb_Plus4));
  Requete1 := getcontroltext('REQUETE1');
  Requete1 := StringReplace(Requete1, '[CRI_DEBMOISANT]', USDateTime(Deb_Moins1), [rfReplaceAll]);
  Requete1 := StringReplace(Requete1, '[CRI_FINMOISANT]', USDateTime(Fin_Moins1), [rfReplaceAll]);
  ExecuteSQL(Requete1);
  Requete2 := getcontroltext('REQUETE2');
  Requete2 := StringReplace(Requete2, '[CRI_DEBMOISANT]', USDateTime(Deb_Moins1), [rfReplaceAll]);
  Requete2 := StringReplace(Requete2, '[CRI_FINMOISANT]', USDateTime(Fin_Moins1), [rfReplaceAll]);
  ExecuteSQL(Requete2);
  Requete3 := getcontroltext('REQUETE3');
  Requete3 := StringReplace(Requete3, '[CRI_DEBMOISANT]', USDateTime(Deb_Moins1), [rfReplaceAll]);
  Requete3 := StringReplace(Requete3, '[CRI_FINMOISANT]', USDateTime(Fin_Moins1), [rfReplaceAll]);
  ExecuteSQL(Requete3);
  Requete4 := getcontroltext('REQUETE4');
  Requete4 := StringReplace(Requete4, '[CRI_DEBMOIS]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete4 := StringReplace(Requete4, '[CRI_FINMOIS]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  ExecuteSQL(Requete4);
  Requete5 := getcontroltext('REQUETE5');
  Requete5 := StringReplace(Requete5, '[CRI_DEBPERSPOST4M]', USDateTime(Deb_Plus4), [rfReplaceAll]);
  Requete5 := StringReplace(Requete5, '[CRI_FINPERSPOST4M]', USDateTime(Fin_Plus4), [rfReplaceAll]);
  ExecuteSQL(Requete5);

end;


procedure TOF_GQR1GII.OnLoad ;
var Stdate,ListeCode,CodeEtat: string;
begin
  Inherited ;
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD") IS NOT NULL) DROP TABLE #PRD ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD1") IS NOT NULL) DROP TABLE #PRD1 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD2") IS NOT NULL) DROP TABLE #PRD2 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD3") IS NOT NULL) DROP TABLE #PRD3 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD4") IS NOT NULL) DROP TABLE #PRD4 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD5") IS NOT NULL) DROP TABLE #PRD5 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD6") IS NOT NULL) DROP TABLE #PRD6 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD7") IS NOT NULL) DROP TABLE #PRD7 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD8") IS NOT NULL) DROP TABLE #PRD8 ');
  ExecuteSQL( ' IF (OBJECT_ID("tempdb..#PRD9") IS NOT NULL) DROP TABLE #PRD9 ');

  Stdate := GetControlText('GPDATEPIECE0');
  if (IsValidDate(StDate)) then Deb_Saisie := DebutDeMois(StrToDate (StDate))
  else  Deb_Saisie := DebutDeMois (V_PGI.DateEntree);

  Stdate := GetControlText('GPDATEPIECE0_');
  if (IsValidDate(StDate)) then Fin_Saisie := FinDeMois(StrToDate(StDate))
  else Fin_Saisie := FinDeMois (V_PGI.DateEntree);
  if (Fin_Saisie < Deb_Saisie) then Fin_Saisie := FinDeMois (Deb_Saisie);

  if  (Fin_Saisie >= V_PGI.DateEntree) then
  begin
    Fin_Moins12 := FinDeMois(PlusMois (Fin_Saisie, -1));
    Deb_Moins12 := DebutDeMois(PlusMois (Fin_Moins12,-11));
    Deb_Plus4 := DebutDeMois ( Fin_Saisie);
    Fin_Plus4 := FinDeMois(PlusMois (Deb_Plus4,3));
  end else
  begin
    Fin_Moins12 := Fin_Saisie ;
    Deb_Moins12 := DebutDeMois(PlusMois (Fin_Moins12,-11));
    Deb_Plus4 := DebutDeMois ( PlusMois (Fin_Saisie,1));
    Fin_Plus4 := FinDeMois(PlusMois (Deb_Plus4,3));
  end;
  // MAJ requête de la fiche
  SetControlText( 'GPDATEPIECE0' , DateToStr (Deb_Saisie));
  SetControlText( 'GPDATEPIECE0_' , DateToStr (Fin_Saisie));
  SetControlText( 'DEBPERSPOST4MS' , DateToStr (Deb_Plus4));

  ListeCode := THMultiValcomboBox(GetControl('CODEMARCHE')).value  ;
  if (trim(ListeCode)<>'') then
  begin
  listeMarche := ' IN ("'+ReadtokenSt(ListeCode)+'"';
  while (trim(ListeCode)<>'') do listeMarche := listeMarche+',"'+ReadtokenSt(ListeCode)+'"';
  listeMarche := listeMarche + ')';
  end else listeMarche := '';

  ListeCode := THMultiValcomboBox(GetControl('CODEAGENCE')).value  ;
  if (trim(ListeCode)<>'') then
  begin
  listeAgence := ' IN ("'+ReadtokenSt(ListeCode)+'"';
  while (trim(ListeCode)<>'') do listeAgence := listeAgence+',"'+ReadtokenSt(ListeCode)+'"';
  listeAgence := listeAgence + ')';     
  end else listeAgence := '';

  CodeEtat := GetControlText ('FETAT') ;
  if  (CodeEtat = 'GIA') then
  RequeteParAgence
  else if (CodeEtat = 'GIM') then
  RequeteParMarche;
end;

procedure TOF_GQR1GII.RequeteParAgence ;
var Requete0,Requete1,Requete2,Requete3,Requete4,Requete5,Requete6 : string;
    Requete7,Requete8,Requete9 : string;
begin
  Requete0 := getcontroltext('REQUETE0');
  Requete0 := CritereAgence (Requete0, ' CC2.YX_code ');
  ExecuteSQL(Requete0);
  Requete1 := getcontroltext('REQUETE1');
  Requete1 := StringReplace(Requete1, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete1 := StringReplace(Requete1, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete1 := CritereMarche(Requete1, ' gp_libretiers1 ');
  ExecuteSQL(Requete1);
  Requete2 := getcontroltext('REQUETE2');
  Requete2 := StringReplace(Requete2, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete2 := StringReplace(Requete2, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete2 := CritereMarche(Requete2, ' gp_libretiers1 ');
  ExecuteSQL(Requete2);
  Requete3 := getcontroltext('REQUETE3');
  Requete3 := StringReplace(Requete3, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete3 := StringReplace(Requete3, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete3 := CritereMarche(Requete3, ' gp_libretiers1 ');
  ExecuteSQL(Requete3);
  Requete4 := getcontroltext('REQUETE4');
  Requete4 := StringReplace(Requete4, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete4 := StringReplace(Requete4, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete4 := CritereMarche(Requete4, ' ytc_tablelibretiers1 ');
  ExecuteSQL(Requete4);
  Requete5 := getcontroltext('REQUETE5');
  Requete5 := StringReplace(Requete5, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete5 := StringReplace(Requete5, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete5 := CritereMarche(Requete5, ' ytc_tablelibretiers1 ');
  ExecuteSQL(Requete5);
  Requete6 := getcontroltext('REQUETE6');
  Requete6 := StringReplace(Requete6, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete6 := StringReplace(Requete6, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete6 := CritereMarche(Requete6, ' ytc_tablelibretiers1 ');
  ExecuteSQL(Requete6);
  Requete7 := getcontroltext('REQUETE7');
  Requete7 := StringReplace(Requete7, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete7 := StringReplace(Requete7, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete7 := CritereMarche(Requete7, ' ytc_tablelibretiers1 ');
  ExecuteSQL(Requete7);
  Requete8 := getcontroltext('REQUETE8');
  Requete8 := StringReplace(Requete8, '[CRI_DEBPERSPOST4M]', USDateTime(Deb_Plus4), [rfReplaceAll]);
  Requete8 := StringReplace(Requete8, '[CRI_FINPERSPOST4M]', USDateTime(Fin_Plus4), [rfReplaceAll]);
  Requete8 := CritereMarche(Requete8, ' ytc_tablelibretiers1 ');
  ExecuteSQL(Requete8);
  Requete9 := getcontroltext('REQUETE9');
  Requete9 := StringReplace(Requete9, '[CRI_DEBPERSANT12M]', USDateTime(Deb_Moins12), [rfReplaceAll]);
  Requete9 := StringReplace(Requete9, '[CRI_FINPERSANT12M]', USDateTime(Fin_Moins12), [rfReplaceAll]);
  Requete9 := CritereMarche(Requete9, ' ytc_tablelibretiers1 ');
  ExecuteSQL(Requete9);
end;

procedure TOF_GQR1GII.RequeteParMarche ;
var Requete0,Requete1,Requete2,Requete3,Requete4,Requete5,Requete6 : string;
    Requete7,Requete8,Requete9 : string;
begin
  Requete0 := getcontroltext('REQUETE0');
  Requete0 := CritereMarche (Requete0, ' CC1.YX_code ');
  ExecuteSQL(Requete0);
  Requete1 := getcontroltext('REQUETE1');
  Requete1 := StringReplace(Requete1, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete1 := StringReplace(Requete1, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete1 := CritereAgence(Requete1, ' gp_libretiers2 ');
  ExecuteSQL(Requete1);
  Requete2 := getcontroltext('REQUETE2');
  Requete2 := StringReplace(Requete2, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete2 := StringReplace(Requete2, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete2 := CritereAgence(Requete2, ' gp_libretiers2 ');
  ExecuteSQL(Requete2);
  Requete3 := getcontroltext('REQUETE3');
  Requete3 := StringReplace(Requete3, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete3 := StringReplace(Requete3, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete3 := CritereAgence(Requete3, ' gp_libretiers2 ');
  ExecuteSQL(Requete3);
  Requete4 := getcontroltext('REQUETE4');
  Requete4 := StringReplace(Requete4, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete4 := StringReplace(Requete4, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete4 := CritereAgence(Requete4, ' ytc_tablelibretiers2 ');
  ExecuteSQL(Requete4);
  Requete5 := getcontroltext('REQUETE5');
  Requete5 := StringReplace(Requete5, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete5 := StringReplace(Requete5, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete5 := CritereAgence(Requete5, ' ytc_tablelibretiers2 ');
  ExecuteSQL(Requete5);
  Requete6 := getcontroltext('REQUETE6');
  Requete6 := StringReplace(Requete6, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete6 := StringReplace(Requete6, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete6 := CritereAgence(Requete6, ' ytc_tablelibretiers2 ');
  ExecuteSQL(Requete6);
  Requete7 := getcontroltext('REQUETE7');
  Requete7 := StringReplace(Requete7, '[CRI_GPDATEPIECE]', USDateTime(Deb_Saisie), [rfReplaceAll]);
  Requete7 := StringReplace(Requete7, '[CRI_GPDATEPIECE_]', USDateTime(Fin_Saisie), [rfReplaceAll]);
  Requete7 := CritereAgence(Requete7, ' ytc_tablelibretiers2 ');
  ExecuteSQL(Requete7);
  Requete8 := getcontroltext('REQUETE8');
  Requete8 := StringReplace(Requete8, '[CRI_DEBPERSPOST4M]', USDateTime(Deb_Plus4), [rfReplaceAll]);
  Requete8 := StringReplace(Requete8, '[CRI_FINPERSPOST4M]', USDateTime(Fin_Plus4), [rfReplaceAll]);
  Requete8 := CritereAgence(Requete8, ' ytc_tablelibretiers2 ');
  ExecuteSQL(Requete8);
  Requete9 := getcontroltext('REQUETE9');
  Requete9 := StringReplace(Requete9, '[CRI_DEBPERSANT12M]', USDateTime(Deb_Moins12), [rfReplaceAll]);
  Requete9 := StringReplace(Requete9, '[CRI_FINPERSANT12M]', USDateTime(Fin_Moins12), [rfReplaceAll]);
  Requete9 := CritereAgence(Requete9, ' ytc_tablelibretiers2 ');
  ExecuteSQL(Requete9);
end;

Function TOF_GQR1GII.CritereAgence(Requete,champ :string):string ;
var stcritere :string;
begin
  if trim(listeAgence) = '' then stcritere := ' 1=1 '
  else stcritere := champ+listeAgence;
  Result :=  StringReplace( Requete,'[CRI_AGENCE]',stcritere,[rfIgnoreCase])
end;

Function TOF_GQR1GII.CritereMarche(Requete,champ :string):string ;
var stcritere :string;
begin
  if trim(listeMarche) = '' then stcritere := ' 1=1 '
  else stcritere := champ+listeMarche;
  Result :=  StringReplace( Requete,'[CRI_MARCHE]',stcritere,[rfIgnoreCase])
end;

Initialization
  registerclasses ( [ TOF_GQR1GTV ] ) ;
  registerclasses ( [ TOF_GQR1GII ] ) ;
  registerclasses ( [ TOF_GQR1GPA ] ) ;
end.
