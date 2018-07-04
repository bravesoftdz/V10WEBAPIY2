{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 31/07/2001
Modifié le ... :   /  /
Description .. : Edition des cumuls
Mots clefs ... : PAIE;CUMUL
*****************************************************************}
{
 PT1 31/07/2001 V547 SB Modification champ suffixe MODEREGLE
 PT2 23/01/2002 V571 SB Suppression du join table CumulPaie
 PT3 28/01/2002 V571 SB Sélection alphanumérique du code salarié
 PT4 18/04/2002 V571 SB Fiche de bug n°10087 : V_PGI_env.LibDossier non renseigné en Mono
 PT5 24/01/2003 V591 SB FQ 10433 contrôle zone cumul
 PT6 24/01/2003 V591 SB FQ 10433 contrôle zone cumul
 PT7 02/10/2003 V_42 SB Affichage des ongles si gestion paramsoc des combos libres
 PT8-1 27/11/2003 V_50 SB FQ 10105 Edition sur 12 mois glissants
 PT8-2 27/11/2003 V_50 SB Suppression variables globales de pgedtetat
 PT8-3 27/11/2003 V_50 SB FQ 10982 Suppression option monnaie inversée
 PT9 11/07/2005 V_60 PH FQ 11337 Controle des bornes de debut et de fin que si les cumuls
                        sont numériques
 PT10 21/09/2005 V_65 SB Traitement des combos libres en rupture
 PT11 20/07/2006 V_65 SB Suite FQ 12813 Contrôle cohérence periode
 PT12 28/06/2007 V_72 FC FQ 13508 Ajout case salariés sortis
 PT13 27/07/2007 V_72 FC FQ 14490 pb préaffichage des dates début et fin quand aucune paie effectuée sur un exercice
 }

unit UTOFPGEditCumul;

interface

uses StdCtrls, Controls, Classes, sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  eQRS1, UTOB,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF,
  ParamDat, ParamSoc, HQry;

type
  TOF_PGFICHECUMUL = class(TOF)
  private
    AddChamp: string;
    procedure DateElipsisclick(Sender: TObject);
    procedure Change(Sender: TObject);
    procedure ChangeLieuTravail(Sender: TObject);
    procedure ChangeExercice(Sender: TObject);
    //procedure MonnaieInverse(Sender: TObject); PT8-3
    function CreateOrderBy: string;
    procedure ControlPeriodeDeb(Sender: TObject);
    procedure ControlPeriodeFin(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    procedure ChangeCumul(Sender: TObject);
    procedure PgIFValidPeriode; { PT8-1 }
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
  end;
implementation

uses PgEditOutils,PgEditOutils2, EntPaie, PGoutils2;

{-------------------------------------------------------------------------------
                     TOF  FICHE CUMUL
--------------------------------------------------------------------------------}
procedure TOF_PGFICHECUMUL.OnArgument(Arguments: string);
var
  DateDeb, DateFin, Defaut: THEdit;
  Exercice: THValComboBox;
  QPeriode, QExer: TQuery;
  CSal, Alpha, CTrav1, CTrav2, CTrav3, CTrav4, CTrav5, CEtab: TCheckBox;
  Min, Max: string;
  Ckeuro, CkPaie, Check: TCheckBox;
  DebPer,FinPer,ExerPerEncours : string;  //PT13
begin
  //PGEDITION:='CUMUL'; { PT8-2 }
  Ckeuro := TCheckBox(GetControl('CKEURO'));
  if Ckeuro <> nil then Ckeuro.Checked := VH_Paie.PGTenueEuro;
  CkPaie := TCheckBox(GetControl('CKPAIE'));
  if CkPaie <> nil then CkPaie.Checked := VH_Paie.PGDecalage;
  VisibiliteChamp(Ecran);
  VisibiliteChampLibre(Ecran);
  { DEB PT7 }
  SetControlProperty('TBCOMPLEMENT', 'Tabvisible', (VH_Paie.PGNbreStatOrg > 0) or (VH_Paie.PGLibCodeStat <> ''));
  SetControlProperty('TBCHAMPLIBRE', 'Tabvisible', (VH_Paie.PgNbCombo > 0));
  { FIN PT7 }
  //Valeur par défaut
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Defaut := ThEdit(getcontrol('PHC_SALARIE'));
  if Defaut <> nil then
  begin
    Defaut.text := Min;
    Defaut.OnExit := ExitEdit;
  end;
  Defaut := ThEdit(getcontrol('PHC_SALARIE_'));
  if Defaut <> nil then
  begin
    Defaut.text := Max;
    Defaut.OnExit := ExitEdit;
  end;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Defaut := ThEdit(getcontrol('PHC_ETABLISSEMENT'));
  if Defaut <> nil then
  begin
    Defaut.text := Min;
  end;
  Defaut := ThEdit(getcontrol('PHC_ETABLISSEMENT_'));
  if Defaut <> nil then
  begin
    Defaut.text := Max;
  end;
  RecupMinMaxTablette('PG', 'CUMULPAIE', 'PCL_CUMULPAIE', Min, Max);
  Defaut := THEdit(GetControl('PHC_CUMULPAIE'));
  if Defaut <> nil then
  begin
    Defaut.MaxLength := 2;
    Defaut.text := Min;
    Defaut.OnExit := ChangeCumul;
  end;
  Defaut := THEdit(GetControl('PHC_CUMULPAIE_'));
  if Defaut <> nil then
  begin
    Defaut.MaxLength := 2;
    Defaut.text := Max;
    Defaut.OnExit := ChangeCumul;
  end;

  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    //  Defaut.text:=V_PGI_env.LibDossier; PT4 Mise en commentaire
    Defaut.text := GetParamSoc('SO_LIBELLE');

  Exercice := THValComboBox(getcontrol('EDEXERSOC'));
  DateDeb := ThEdit(getcontrol('XX_VARIABLEDEB'));
  DateFin := ThEdit(getcontrol('XX_VARIABLEFIN'));

  QExer := OpenSql('SELECT MAX(PEX_EXERCICE),MAX(PEX_DATEDEBUT) FROM EXERSOCIAL ' +
    'WHERE PEX_ACTIF="X"', true);
  if not QExer.EOF then //PORTAGECWAS
    if Exercice <> nil then Exercice.value := QExer.Fields[0].asstring;
  Ferme(QExer);

  QPeriode := OpenSql('SELECT MIN(PPU_DATEDEBUT),MAX(PPU_DATEFIN) FROM PAIEENCOURS ' +
    'LEFT JOIN EXERSOCIAL ON PPU_DATEDEBUT>=PEX_DATEDEBUT AND PPU_DATEFIN<=PEX_DATEFIN WHERE PEX_EXERCICE="' + Exercice.value + '"', TRUE);
  if (not QPeriode.EOF) and (QPeriode.Fields[0].AsDateTime <> 0) then //PT13
  begin
    SetControlText('XX_VARIABLEDEB', DateToStr(QPeriode.Fields[0].AsDateTime));
    SetControlText('XX_VARIABLEFIN', DateToStr(QPeriode.Fields[1].AsDateTime));
  end
  else
  begin
    //DEB PT13
    ExerPerEncours := Exercice.value;
    if RendPeriodeEnCours(ExerPerEncours,DebPer,FinPer) = True then
    begin
      SetControlText('XX_VARIABLEDEB', DebPer);
      SetControlText('XX_VARIABLEFIN', FinPer);
    end;
//    SetControlText('XX_VARIABLEDEB', DateToStr(idate1900));
//    SetControlText('XX_VARIABLEFIN', DateToStr(idate1900));
    //FIN PT13
  end;
  Ferme(QPeriode);

  if (DateDeb <> nil) and (DateFin <> nil) and (Exercice <> nil) then
  begin
    DateDeb.OnElipsisClick := DateElipsisclick;
    DateDeb.OnExit := ControlPeriodeDeb;
    DateFin.OnElipsisClick := DateElipsisclick;
    DateFin.OnExit := ControlPeriodeFin;
    Exercice.Onchange := ChangeExercice;
  end;

  CTrav1 := TCheckBox(GetControl('CN1'));
  if CTrav1 <> nil then CTrav1.OnClick := ChangeLieuTravail;
  CTrav2 := TCheckBox(GetControl('CN2'));
  if CTrav2 <> nil then CTrav2.OnClick := ChangeLieuTravail;
  CTrav3 := TCheckBox(GetControl('CN3'));
  if CTrav3 <> nil then CTrav3.OnClick := ChangeLieuTravail;
  CTrav4 := TCheckBox(GetControl('CN4'));
  if CTrav4 <> nil then CTrav4.OnClick := ChangeLieuTravail;
  CTrav5 := TCheckBox(GetControl('CN5'));
  if CTrav5 <> nil then CTrav5.OnClick := ChangeLieuTravail;
  CEtab := TCheckBox(GetControl('CETAB'));
  if CEtab <> nil then CEtab.OnClick := ChangeLieuTravail;
  CSal := TCheckBox(GetControl('CSAL'));
  if CSal <> nil then
  begin
    CSal.Checked := True;
    CSal.OnClick := Change;
  end;
  Alpha := TCheckBox(GetControl('CALPHA'));
  if Alpha <> nil then Alpha.OnClick := Change;
  Check := TCheckBox(GetControl('CL1'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL2'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL3'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL4'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;


  Defaut := THEdit(GetControl('XX_RUPTURE2'));
  if Defaut <> nil then Defaut.text := 'PHC_SALARIE';

  {PT8-3 ChMonInv := TCheckBox(GetControl('CHMONNAIEINV'));
  if ChMonInv<>nil then begin ChMonInv.OnClick:=MonnaieInverse; PgMonnaieInv:=ChMonInv.checked; End;}
end;


procedure TOF_PGFICHECUMUL.Change(Sender: TObject);
var
  Edit: THEdit;
  CSal, CAlpha: TCheckBox;
begin
  //Affectation champs invisible
  Edit := THEdit(GetControl('XX_RUPTURE2'));
  CSal := TCheckBox(GetControl('CSAL'));
  CAlpha := TCheckBox(GetControl('CALPHA'));
  if (CSal <> nil) and (CAlpha <> nil) and (Edit <> nil) then
  begin
    if CSal.Checked = True then
    begin
      CAlpha.enabled := True;
      Edit.Text := 'PHC_SALARIE';
      SetControlEnabled('CKSORTIE',True);   //PT12
    end
    else
      if CSal.Checked = False then
    begin
      CAlpha.Checked := False;
      CAlpha.enabled := False;
      Edit.Text := '';
      SetControlEnabled('CKSORTIE',False);   //PT12
      SetControlChecked('CKSORTIE', False);  //PT12
    end;
  end;
  if TCheckBox(Sender).Name = 'CALPHA' then //PT3
    AffectCritereAlpha(Ecran, TCheckBox(Sender).Checked, 'PHC_SALARIE', 'LIBELLE');
end;

procedure TOF_PGFICHECUMUL.ChangeExercice(Sender: TObject);
var
  QPeriode: TQuery;
  DebPer,FinPer,ExerPerEncours : string;  //PT13
begin
  QPeriode := OpenSql('SELECT MIN(PPU_DATEDEBUT),MAX(PPU_DATEFIN) FROM PAIEENCOURS ' +
    'LEFT JOIN EXERSOCIAL ON PPU_DATEDEBUT>=PEX_DATEDEBUT AND PPU_DATEFIN<=PEX_DATEFIN WHERE PEX_EXERCICE="' + GetControlText('EDEXERSOC') + '"', TRUE);
  if (not QPeriode.Eof) and (QPeriode.Fields[0].AsDateTime <> 0) then  //PT13
  begin
    SetControlText('XX_VARIABLEDEB', DateToStr(QPeriode.Fields[0].AsDateTime));
    SetControlText('XX_VARIABLEFIN', DateToStr(QPeriode.Fields[1].AsDateTime));
    PgIFValidPeriode;  { PT11 }
  end
  else
  begin
    //DEB PT13
    ExerPerEncours := GetControlText('EDEXERSOC');
    if RendPeriodeEnCours(ExerPerEncours,DebPer,FinPer) = True then
    begin
      SetControlText('XX_VARIABLEDEB',DebPer);
      SetControlText('XX_VARIABLEFIN',FinPer);
    end;
    //FIN PT13
  end;
  Ferme(QPeriode);
end;

function TOF_PGFICHECUMUL.CreateOrderBy: string;
var
  Ck,CSal, CAlpha : TCheckBox;
  StRupt, StAlpha, StSal, St: string;
begin
  StRupt := '';
  StAlpha := '';
  StSal := '';
  St := '';
  AddChamp := '';
  CSal := TCheckBox(GetControl('CSAL'));
  CAlpha := TCheckBox(GetControl('CALPHA'));
  Ck := TCheckBox(GetControl('CN1'));
  if Ck <> nil then if Ck.Checked = True then StRupt := 'PHC_TRAVAILN1';
  Ck := TCheckBox(GetControl('CN2'));
  if Ck <> nil then if Ck.Checked = True then StRupt := 'PHC_TRAVAILN2';
  Ck := TCheckBox(GetControl('CN3'));
  if Ck <> nil then if Ck.Checked = True then StRupt := 'PHC_TRAVAILN3';
  Ck := TCheckBox(GetControl('CN4'));
  if Ck <> nil then if Ck.Checked = True then StRupt := 'PHC_TRAVAILN4';
  Ck := TCheckBox(GetControl('CN5'));
  if Ck <> nil then if Ck.Checked = True then StRupt := 'PHC_CODESTAT';
  Ck := TCheckBox(GetControl('CETAB'));
  if Ck <> nil then if Ck.Checked = True then StRupt := 'PHC_ETABLISSEMENT';

  { DEB PT10 }
  Ck := TCheckBox(GetControl('CL1'));
  if Ck <> nil then if Ck.Checked = True then StRupt := 'PHC_LIBREPCMB1';
  Ck := TCheckBox(GetControl('CL2'));
  if Ck <> nil then if Ck.Checked = True then StRupt := 'PHC_LIBREPCMB2';
  Ck := TCheckBox(GetControl('CL3'));
  if Ck <> nil then if Ck.Checked = True then StRupt := 'PHC_LIBREPCMB3';
  Ck := TCheckBox(GetControl('CL4'));
  if Ck <> nil then if Ck.Checked = True then StRupt := 'PHC_LIBREPCMB4';
  { FIN PT10 }


  //PgChampRupt:=StRupt; { PT8-2 }
  if StRupt <> '' then AddChamp := StRupt + ',';
  if CSal <> nil then if CSal.Checked = True then StSal := 'PHC_SALARIE';
  if CAlpha <> nil then if CAlpha.Checked = True then StAlpha := 'PSA_LIBELLE';

  if (StRupt <> '') and (StSal = '') and (StAlpha = '') then St := StRupt;
  if (StRupt = '') and (StSal <> '') and (StAlpha = '') then St := StSal;
  if (StRupt = '') and (StSal = '') and (StAlpha <> '') then St := StAlpha;

  if (StRupt <> '') and (StSal <> '') and (StAlpha <> '') then St := StRupt + ',' + StAlpha + ',' + StSal;

  if (StRupt <> '') and (StSal <> '') and (StAlpha = '') then St := StRupt + ',' + StSal;
  if (StRupt = '') and (StSal <> '') and (StAlpha <> '') then St := StAlpha + ',' + StSal;
  if (StRupt <> '') and (StSal = '') and (StAlpha <> '') then St := StRupt + ',' + StAlpha;

  if St <> '' then result := St + ',PHC_CUMULPAIE';
  if st = '' then result := 'PHC_CUMULPAIE';
end;

procedure TOF_PGFICHECUMUL.OnUpdate;
var
  Temp, Tempo, Critere: string;
  Pages: TPageControl;
  x: integer;
  DateDeb, DateFin: ThEdit;
  Order: string;
  CSal: TCheckBox;
begin
  //PGSalarie:=False; { PT8-2 }
  CSal := TCheckBox(GetControl('CSAL'));
  //if CSal<>nil then  PGSalarie:=CSal.Checked; { PT8-2 }

  Pages := TPageControl(GetControl('Pages'));
  Temp := RecupWhereCritere(Pages);
  tempo := '';
  critere := '';
  x := Pos('(', Temp);
  if x > 0 then Tempo := copy(Temp, x, (Length(temp) - 5));
  if tempo <> '' then critere := 'AND ' + Tempo;

  DateDeb := THEdit(GetControl('XX_VARIABLEDEB'));
  DateFin := THEdit(GetControl('XX_VARIABLEFIN'));
  { DEB PT12 }
  if (GetControltext('CKSORTIE')='X') and (UsdateTime(StrToDate(GetControlText('XX_VARIABLEFIN')))<>UsdateTime(idate1900)) then
  Critere:=Critere+' AND PHC_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_DATESORTIE IS NULL '+
                      'OR PSA_DATESORTIE="'+Usdatetime(idate1900)+'" '+
                      'OR PSA_DATESORTIE>"'+UsdateTime(StrToDate(GetControlText('XX_VARIABLEFIN')))+'")';
  { FIN PT12 }

  Order := CreateOrderBy;
  //PGCritere:=Critere; { PT8-2 }
  SetControlText('STWHERE', Critere);
  if GetControlText('CALPHA') = 'X' then //DEB PT3
  begin
    if GetControlText('LIBELLE') <> '' then Critere := Critere + ' AND PSA_LIBELLE>="' + GetControlText('LIBELLE') + '" ';
    if GetControlText('LIBELLE_') <> '' then Critere := Critere + ' AND PSA_LIBELLE<="' + GetControlText('LIBELLE_') + '" ';
  end; //FIN PT3

  if CSal <> nil then
    if CSal.Checked = True then
    begin
      TFQRS1(Ecran).WhereSQL := 'SELECT DISTINCT PHC_SALARIE,PHC_CUMULPAIE,' + {PCL_LIBELLE,  //PT2
      'PCL_PREDEFINI,'+ }
      'PSA_SALARIE,PSA_LIBELLE,PSA_NUMEROSS,PSA_PRENOM,PSA_ADRESSE1,PSA_ADRESSE3,' +
        'PSA_ADRESSE2,PSA_CODEPOSTAL,PSA_VILLE,PSA_DATEENTREE,PSA_DATESORTIE,' + AddChamp + ' ' +
        'PSA_QUALIFICATION,PSA_COEFFICIENT,PSA_CODEEMPLOI,PSA_PGMODEREGLE,PSA_LIBELLEEMPLOI ' + {PT1}
      'FROM HISTOCUMSAL ' +
        'LEFT JOIN SALARIES ON PSA_SALARIE=PHC_SALARIE ' +
        //  'LEFT JOIN CUMULPAIE ON PHC_CUMULPAIE=PCL_CUMULPAIE AND ##PCL_PREDEFINI##'+'PCL_CUMULPAIE <> "ZZZ" '+ PT2
      'WHERE PHC_DATEDEBUT>="' + UsDateTime(StrToDate(DateDeb.text)) + '"  ' +
        'AND PHC_DATEFIN<="' + UsDateTime(StrToDate(DateFin.text)) + '" ' +
        '' + critere + ' ' +
        'ORDER BY ' + Order + '';
    end;
  if CSal.Checked = False then
  begin
    TFQRS1(Ecran).WhereSQL := 'SELECT DISTINCT ' + AddChamp + 'PHC_CUMULPAIE ' + {PCL_LIBELLE,'+ //PT2
    'PCL_PREDEFINI '+ }
    'FROM HISTOCUMSAL ' +
      //'LEFT JOIN CUMULPAIE ON PHC_CUMULPAIE=PCL_CUMULPAIE AND ##PCL_PREDEFINI##'+'PCL_CUMULPAIE <> "ZZZ" '+ PT2
    'WHERE PHC_DATEDEBUT>="' + UsDateTime(StrToDate(DateDeb.text)) + '"  ' +
      'AND PHC_DATEFIN<="' + UsDateTime(StrToDate(DateFin.text)) + '" ' +
      '' + critere + ' ' +
      'ORDER BY ' + Order + '';
  end;

end;
{ PT8-3 Mise en commentaire
procedure TOF_PGFICHECUMUL.MonnaieInverse(Sender: TObject);
var
ChMonInv : TCheckBox;
begin
ChMonInv := TCheckBox(GetControl('CHMONNAIEINV'));
if (ChMonInv<>nil)  then  PgMonnaieInv:=ChMonInv.checked;
PgTauxConvert:=RendTauxConvertion;
end;                        }

procedure TOF_PGFICHECUMUL.ChangeLieuTravail(Sender: TObject);
begin
  RecupChampRupture(Ecran);
  BloqueChampLibre(Ecran);
end;

procedure TOF_PGFICHECUMUL.ControlPeriodeDeb(Sender: TObject);
begin
  PgIFValidPeriode;
  { PT8-1 Mise en commentaire
  Q := OpenSql('SELECT PEX_DATEDEBUT FROM EXERSOCIAL WHERE PEX_EXERCICE="'+GetControlText('EDEXERSOC')+'" ', True);
  if not Q.eof then
    DebExer:=Q.findfield('PEX_DATEDEBUT').AsDateTime
  Else
    DebExer:=Idate1900;
  Ferme(Q);
  if (StrToDate(GetControlText('XX_VARIABLEDEB'))<DebExer) and (DebExer<>idate1900) then
     Begin
     PgiBox('La date de début ne peut être inferieur au '+DateToStr(DebExer)+'','Date Erronée!');
     SetControlText('XX_VARIABLEDEB',DateToStr(DebExer));
     End;}
end;

procedure TOF_PGFICHECUMUL.ControlPeriodeFin(Sender: TObject);
begin
  PgIFValidPeriode;
  {PT8-1 Mise en commentaire
  Q := OpenSql('SELECT PEX_DATEFIN FROM EXERSOCIAL WHERE PEX_EXERCICE="'+GetControlText('EDEXERSOC')+'" ', True);
  if Not Q.Eof then    //PORTAGECWAS
    FinExer:=Q.findfield('PEX_DATEFIN').AsDateTime
  Else
    FinExer:=idate1900;
  Ferme(Q);
  if (StrToDate(GetControlText('XX_VARIABLEFIN'))>FinExer) and (FinExer<>idate1900) then
     Begin
     PgiBox('La date de fin ne peut être supérieur au '+DateToStr(FinExer)+'','Date Erronée!');
     SetControlText('XX_VARIABLEFIN',DateToStr(FinExer));
     End;}
end;

procedure TOF_PGFICHECUMUL.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if edit.text <> '' then
      if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit, 10);
end;


procedure TOF_PGFICHECUMUL.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGFICHECUMUL.ChangeCumul(Sender: TObject);
var
  Code, CodeCum, CodeCum_: integer;
begin
  { DEB PT6 }
  SetControlText('PHC_CUMULPAIE', AffectDefautCode(THEdit(GetControl('PHC_CUMULPAIE')), 2));
  SetControlText('PHC_CUMULPAIE_', AffectDefautCode(THEdit(GetControl('PHC_CUMULPAIE_')), 2));
  { DEB PT6 }
  { DEB PT5 }
  if (GetControlText('PHC_CUMULPAIE') <> '') and (GetControlText('PHC_CUMULPAIE_') <> '') then
  begin
    val(GetControlText('PHC_CUMULPAIE'), CodeCum, code);
    val(GetControlText('PHC_CUMULPAIE_'), CodeCum_, code);
    if ((CodeCum) > (CodeCum_)) and (IsNumeric (GetControlText('PHC_CUMULPAIE')) and IsNumeric (GetControlText('PHC_CUMULPAIE_'))) then // PT9
      PgiBox('La borne de fin doit être supérieure à la borne de début.', TFQRS1(Ecran).caption);
  end;
  { FIN PT5 }
end;
{ DEB PT8-1 }
procedure TOF_PGFICHECUMUL.PgIFValidPeriode;
var
  YYD, MMD, JJ, YYF, MMF: WORD;
begin
  if IsValidDate(GetControlText('XX_VARIABLEDEB')) and IsValidDate(GetControlText('XX_VARIABLEFIN')) then
  begin
    DecodeDate(StrToDate(GetControlText('XX_VARIABLEDEB')), YYD, MMD, JJ);
    DecodeDate(StrToDate(GetControlText('XX_VARIABLEFIN')), YYF, MMF, JJ);
    if (YYF > YYD) and (MMF >= MMD) then
    begin
      PgiBox('La période d''édition ne peut excéder douze mois civils.', 'Date Erronée!'); { PT8-1 04/05/2004 }
      SetControlText('XX_VARIABLEFIN', DateToStr(FinDeMois(PlusDate(StrToDate(GetControlText('XX_VARIABLEDEB')), 11, 'M'))));
    end;
  end;
end;
{ FIN PT8-1 }

initialization
  registerclasses([TOF_PGFICHECUMUL]);
end.

