{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 31/07/2001
Modifié le ... :   /  /
Description .. : Edition de la fiche individuelle
Mots clefs ... : PAIE;FICHE INDIVIDUELLE
*****************************************************************
 PT1 31/07/2001 V547 SB Modification champ suffixe MODEREGLE
 PT2 07/12/2001 V563 SB Fiche de bug n°355 Les montants additionnés =0 ne s'imprime pas
 PT3-1 24/01/2002 V571 SB Fiche de Bug n°441 Rubrique edité plusieurs fois
 PT3-2 24/01/2002 V571 SB Utilisation de la fiche et de la tof du menu pour l'edition
                          de la fiche ind.
                          FICHEINDSAL_ETAT Et TOF_PGFICHEINDBUL ne st plus utilisé
 PT4 28/01/2002 V571 SB Sélection alphanumérique du code salarié
 PT5 25/03/2002 V571 SB Affectation des periodes de paie en edition saisie bulletin
 PT6 18/04/2002 V571 SB Fiche de bug n°10087 : V_PGI_env.LibDossier non renseigné en Mono
 PT7 02/10/2002 V585 SB FQ n°10214 Bulletin bimensuel non edité si selection mensualité de fin
 PT8 29/09/2003 V_42 SB FQ n°10694 Edition exercice précedant les dates de paie via saisie de la paie
 PT9 02/10/2003 V_42 SB Affichage des ongles si gestion paramsoc des combos libres
 PT10-1 27/11/2003 V_50 SB FQ 10104 Edition sur 12 mois glissants
 PT10-2 27/11/2003 V_50 SB Suppression variables globales de pgedtetat
 PT10-3 27/11/2003 V_50 SB FQ 10982 Suppression option monnaie inversée
 PT11   11/12/2003 V_50 SB FQ 10245 Ajout option rubrique non imprimable
 PT12   30/09/2004 V_50 JL FQ 11613 Gestion salariés confidentiel
 PT13   14/12/2004 V_60 JL FQ 11784 Gestion salariés confidentiel pour recap
                           + Correction si recap le 21/01/2004 HISTOBULLETIN au lieu de SALARIES
 PT14   18/03/2004 V_60 SB FQ 11398 Ajout case à coher pour édition des rubriques patronales
 PT15   01/07/2005 V_65 SB FQ 12300 Gestion raccourci de l'attestation maladie
 PT16   21/09/2005 V_65 SB Traitement des combos libres en rupture
 PT17   22/11/2005 V_65 SB Ajout case salariés sortis
 PT18   29/03/2007 V_72 FC FQ 13892 pb préaffichage des dates début et fin quand aucune paie effectuée sur un exercice
 PT19   28/05/2007 V_72 JL Ajout accès depuis la fiche salarié
 }
unit UTOFPGEditFicInd;

interface
uses StdCtrls, Controls, Classes,sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  eQRS1, UTOB,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, ParamDat,
  ParamSoc, HQry;

type

  TOF_PGFICHEIND = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
  private
    AddChamp, ExDateDeb, ExDateFin: string;
    Recap: Boolean;
    procedure DateElipsisclick(Sender: TObject);
    procedure Change(Sender: TObject);
    procedure ChangeLieuTravail(Sender: TObject);
    procedure ChangeExercice(Sender: TObject);
    //procedure MonnaieInverse(Sender: TObject);  PT10-3
    function  CreateOrderBy: string;
    procedure ControlPeriodeDeb(Sender: TObject);
    procedure ControlPeriodeFin(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    procedure PgIFValidPeriode;
  end;
implementation

uses PgEditOutils,PgEditOutils2, EntPaie,PGoutils2;

{-------------------------------------------------------------------------------
                     TOF  FICHE INDIVIDUEL
--------------------------------------------------------------------------------}
procedure TOF_PGFICHEIND.OnArgument(Arguments: string);
var
  THDateDeb, THDateFin, Defaut: THEdit;
  Exercice: THValComboBox;
  QPeriode, QExer: TQuery;
  //CSal, Alpha, CTrav1, CTrav2, CTrav3, CTrav4, CTrav5, CEtab, ChMonInv: TCheckBox;
  Min, Max, st, Origine, CodeSal: string;
  check: TCheckBox;
  DebPer,FinPer,ExerPerEncours : string;  //PT18
begin
  Recap := False;
  //PGTypeFiche:='SALARIE'; PT10-2
  //PGEDITION:='IND';
  //SetControlText('DOSSIER',V_PGI_env.LibDossier);     //PT6 Mise en commentaire
  SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
  SetControlChecked('CKEURO', VH_Paie.PGTenueEuro);
  SetControlChecked('CKPAIE', VH_Paie.PGDecalage);
  SetControlText('XX_RUPTURE2', 'PHB_SALARIE');

  QExer := OpenSql('SELECT MAX(PEX_EXERCICE),MAX(PEX_DATEDEBUT) FROM EXERSOCIAL ' +
    'WHERE PEX_ACTIF="X"', true);
  if not QExer.eof then
    SetControlText('EDEXERSOC', QExer.Fields[0].asstring)
  else
    SetControlText('EDEXERSOC', '');
  Ferme(QExer);

  QPeriode := OpenSql('SELECT MIN(PPU_DATEDEBUT),MAX(PPU_DATEFIN) FROM PAIEENCOURS ' +
    'LEFT JOIN EXERSOCIAL ON PPU_DATEDEBUT>=PEX_DATEDEBUT AND PPU_DATEFIN<=PEX_DATEFIN ' +
    'WHERE PEX_EXERCICE="' + GetControltext('EDEXERSOC') + '"', TRUE);
  if (not QPeriode.eof) and (QPeriode.Fields[0].AsDateTime <> 0) then //PT18
  begin
    ExDateDeb := DateToStr(QPeriode.Fields[0].AsDateTime);
    ExDateFin := DateToStr(QPeriode.Fields[1].AsDateTime);
  end
  else
  begin
    //DEB PT18
    ExerPerEncours := GetControlText('EDEXERSOC');
    if RendPeriodeEnCours(ExerPerEncours,DebPer,FinPer) = True then
    begin
      ExDateDeb := DebPer;
      ExDateFin := FinPer;
    end;
    //ExDateDeb := DateToStr(idate1900);
    //ExDateFin := DateToStr(idate1900);
    //FIN PT18
  end;
  Ferme(QPeriode);
  Exercice := THValComboBox(getcontrol('EDEXERSOC'));
  THDateDeb := ThEdit(getcontrol('XX_VARIABLEDEB'));
  THDateFin := ThEdit(getcontrol('XX_VARIABLEFIN'));
  if (THDateDeb <> nil) and (THDateFin <> nil) and (Exercice <> nil) then
  begin
    THDateDeb.text := ExDateDeb;
    THDateDeb.OnElipsisClick := DateElipsisclick;
    THDateDeb.OnExit := ControlPeriodeDeb;
    THDateFin.text := ExDateFin;
    THDateFin.OnElipsisClick := DateElipsisclick;
    THDateFin.OnExit := ControlPeriodeFin;
    Exercice.OnChange := ChangeExercice;
  end;

  st := Trim(Arguments);
  Origine := ReadTokenSt(st);
  {DEB PT15 Dans le cas de l'appel via l'attestation maladie }
  if Origine = 'MAL' then
  begin
    CodeSal := ReadTokenSt(st); // Recup code Salarie
    SetControlText('PHB_SALARIE', CodeSal);
    SetControlText('PHB_SALARIE_', CodeSal);
    SetControlChecked('CSAL', True);
    ChangeExercice(nil);
    TFQRS1(Ecran).bAgrandir.Visible := False;
    TFQRS1(Ecran).bAgrandir.Click;
    TFQRS1(Ecran).HPB.Visible := False;
    TFQRS1(Ecran).BValider.Click;
  end
  {FIN PT15 }
  else
  begin
    VisibiliteChamp(Ecran);
    VisibiliteChampLibre(Ecran);
    { DEB PT9 }
    SetControlProperty('TBCOMPLEMENT', 'Tabvisible', (VH_Paie.PGNbreStatOrg > 0) or (VH_Paie.PGLibCodeStat <> ''));
    SetControlProperty('TBCHAMPLIBRE', 'Tabvisible', (VH_Paie.PgNbCombo > 0));
    { FIN PT9 }
    //Valeur par défaut
    RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
    Defaut := ThEdit(getcontrol('PHB_SALARIE'));
    if Defaut <> nil then
    begin
      Defaut.text := Min;
      Defaut.OnExit := ExitEdit;
    end;
    Defaut := ThEdit(getcontrol('PHB_SALARIE_'));
    if Defaut <> nil then
    begin
      Defaut.text := Max;
      Defaut.OnExit := ExitEdit;
    end;
        // DEBUT PT19
    If Origine = 'SAL' then
    begin
      CodeSal := ReadTokenSt(st); // Recup code Salarie
      SetControlText('PHB_SALARIE', CodeSal);
      SetControlText('PHB_SALARIE_', CodeSal);
      SetControlEnabled('PHB_SALARIE', False);
      SetControlEnabled('PHB_SALARIE_', False);
    end;
    // FIN PT19
    RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
    Defaut := ThEdit(getcontrol('PHB_ETABLISSEMENT'));
    if Defaut <> nil then
    begin
      Defaut.text := Min;
    end;
    Defaut := ThEdit(getcontrol('PHB_ETABLISSEMENT_'));
    if Defaut <> nil then
    begin
      Defaut.text := Max;
    end;

    Check := TCheckBox(GetControl('CN1'));
    if Check <> nil then Check.OnClick := ChangeLieuTravail;
    Check := TCheckBox(GetControl('CN2'));
    if Check <> nil then Check.OnClick := ChangeLieuTravail;
    Check := TCheckBox(GetControl('CN3'));
    if Check <> nil then Check.OnClick := ChangeLieuTravail;
    Check := TCheckBox(GetControl('CN4'));
    if Check <> nil then Check.OnClick := ChangeLieuTravail;
    Check := TCheckBox(GetControl('CN5'));
    if Check <> nil then Check.OnClick := ChangeLieuTravail;
    Check := TCheckBox(GetControl('CETAB'));
    if Check <> nil then Check.OnClick := ChangeLieuTravail;
    Check := TCheckBox(GetControl('CL1'));
    if Check <> nil then Check.OnClick := ChangeLieuTravail;
    Check := TCheckBox(GetControl('CL2'));
    if Check <> nil then Check.OnClick := ChangeLieuTravail;
    Check := TCheckBox(GetControl('CL3'));
    if Check <> nil then Check.OnClick := ChangeLieuTravail;
    Check := TCheckBox(GetControl('CL4'));
    if Check <> nil then Check.OnClick := ChangeLieuTravail;

    Check := TCheckBox(GetControl('CSAL'));
    if Check <> nil then
    begin
      Check.Checked := True;
      Check.OnClick := Change;
    end;
    Check := TCheckBox(GetControl('CALPHA'));
    if Check <> nil then Check.OnClick := Change;

    Check := TCheckBox(GetControl('CKPAT'));    { PT14 }
    if Check <> nil then  Check.OnClick := Change;    


    { PT10-3 ChMonInv := TCheckBox(GetControl('CHMONNAIEINV'));
      if ChMonInv<>nil then begin ChMonInv.OnClick:=MonnaieInverse; PgMonnaieInv:=ChMonInv.checked; End;}
  end;
end;

procedure TOF_PGFICHEIND.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGFICHEIND.Change(Sender: TObject);
begin
  //Affectation champs invisible
  if GetControltext('CSAL') = 'X' then
  begin
    Recap := False;
    //PGTypeFiche:='SALARIE'; PT10-2
    SetControlEnabled('CALPHA', True);
    SetControlText('XX_RUPTURE2', 'PHB_SALARIE');
    SetControlEnabled('CKSORTIE',True);   { PT17 }
  end
  else
  begin
    Recap := True;
    //PGTypeFiche:='RECAP';  PT10-2
    SetControlChecked('CALPHA', False);
    SetControlEnabled('CALPHA', False);
    SetControlText('XX_RUPTURE2', '');
    SetControlEnabled('CKSORTIE',False);   { PT17 }
    SetControlChecked('CKSORTIE', False);  { PT17 }
  end;

  if TCheckBox(Sender).Name = 'CALPHA' then //PT4
    AffectCritereAlpha(Ecran, TCheckBox(Sender).Checked, 'PHB_SALARIE', 'PSA_LIBELLE');

  if TCheckBox(Sender).Name = 'CKPAT' then { PT14 }
       Begin
       SetControlEnabled('CKSSRUBPAT',(GetControlText('CKPAT')='-'));
       if GetControlText('CKPAT')='X' then SetControlChecked('CKSSRUBPAT',False);
       End;

end;

procedure TOF_PGFICHEIND.ChangeExercice(Sender: TObject);
var
  QPeriode: TQuery;
  DebPer,FinPer,ExerPerEncours : string;  //PT18
begin
  QPeriode := OpenSql('SELECT MIN(PPU_DATEDEBUT),MAX(PPU_DATEFIN) FROM PAIEENCOURS ' +
    'LEFT JOIN EXERSOCIAL ON PPU_DATEDEBUT>=PEX_DATEDEBUT AND PPU_DATEFIN<=PEX_DATEFIN ' +
    'WHERE PEX_EXERCICE="' + GetControlText('EDEXERSOC') + '"', TRUE);
  if (not QPeriode.eof) and (QPeriode.Fields[0].AsDateTime <> 0) then  //PT18
  begin
    SetControlText('XX_VARIABLEDEB', DateToStr(QPeriode.Fields[0].AsDateTime));
    SetControlText('XX_VARIABLEFIN', DateToStr(QPeriode.Fields[1].AsDateTime));
  end
  else
  begin
    //DEB PT18
    ExerPerEncours := GetControlText('EDEXERSOC');
    if RendPeriodeEnCours(ExerPerEncours,DebPer,FinPer) = True then
    begin
      SetControlText('XX_VARIABLEDEB',DebPer);
      SetControlText('XX_VARIABLEFIN',FinPer);
    end;
    //SetControlText('XX_VARIABLEDEB', DateToStr(idate1900));
    //SetControlText('XX_VARIABLEFIN', DateToStr(idate1900));
    //FIN PT18
  end;
  Ferme(QPeriode);
end;

function TOF_PGFICHEIND.CreateOrderBy: string;
var
  StRupt, StAlpha, StSal, St: string;
begin
  StRupt := '';
  StAlpha := '';
  StSal := '';
  St := '';
  AddChamp := '';
  if GetControlText('CN1') = 'X' then StRupt := 'PHB_TRAVAILN1';
  if GetControlText('CN2') = 'X' then StRupt := 'PHB_TRAVAILN2';
  if GetControlText('CN3') = 'X' then StRupt := 'PHB_TRAVAILN3';
  if GetControlText('CN4') = 'X' then StRupt := 'PHB_TRAVAILN4';
  if GetControlText('CN5') = 'X' then StRupt := 'PHB_CODESTAT';
  if GetControlText('CETAB') = 'X' then StRupt := 'PHB_ETABLISSEMENT';
  { DEB PT16 }
  if GetControlText('CL1') = 'X' then StRupt := 'PHB_LIBREPCMB1';
  if GetControlText('CL2') = 'X' then StRupt := 'PHB_LIBREPCMB2';
  if GetControlText('CL3') = 'X' then StRupt := 'PHB_LIBREPCMB3';
  if GetControlText('CL4') = 'X' then StRupt := 'PHB_LIBREPCMB4';
  { FIN PT16 }
  //PgChampRupt:=StRupt; PT10-2
  if StRupt <> '' then AddChamp := StRupt + ',';
  if GetControlText('CSAL') = 'X' then StSal := 'PHB_SALARIE';
  if GetControlText('CALPHA') = 'X' then StAlpha := 'PSA_LIBELLE';

  if (StRupt <> '') and (StSal = '') and (StAlpha = '') then St := StRupt;
  if (StRupt = '') and (StSal <> '') and (StAlpha = '') then St := StSal;
  if (StRupt = '') and (StSal = '') and (StAlpha <> '') then St := StAlpha;

  if (StRupt <> '') and (StSal <> '') and (StAlpha <> '') then St := StRupt + ',' + StAlpha + ',' + StSal;

  if (StRupt <> '') and (StSal <> '') and (StAlpha = '') then St := StRupt + ',' + StSal;
  if (StRupt = '') and (StSal <> '') and (StAlpha <> '') then St := StAlpha + ',' + StSal;
  if (StRupt <> '') and (StSal = '') and (StAlpha <> '') then St := StRupt + ',' + StAlpha;

  if St <> '' then result := St + ',PHB_ORDREETAT,PHB_NATURERUB,PHB_RUBRIQUE';
  if st = '' then result := 'PHB_ORDREETAT,PHB_NATURERUB,PHB_RUBRIQUE';
end;

procedure TOF_PGFICHEIND.OnUpdate;
var
  Temp, Tempo, Critere, StRubimpr, StRubPat : string;
  Pages: TPageControl;
  x: integer;
  DateDeb, DateFin, Order, StConf : string;
begin
  Pages := TPageControl(GetControl('Pages'));
  Temp := RecupWhereCritere(Pages);
  tempo := '';
  critere := '';
  x := Pos('(', Temp);
  if x > 0 then Tempo := copy(Temp, x, (Length(temp) - 5));
  if tempo <> '' then critere := 'AND ' + Tempo;


  DateDeb := GetControlText('XX_VARIABLEDEB');
  DateFin := GetControlText('XX_VARIABLEFIN');
  { DEB PT17 }
  if (GetControltext('CKSORTIE')='X') and (UsdateTime(StrToDate(DateFin))<>UsdateTime(idate1900)) then
  Critere:=Critere+' AND (PSA_DATESORTIE IS NULL '+
                      'OR PSA_DATESORTIE="'+Usdatetime(idate1900)+'" '+
                      'OR PSA_DATESORTIE>"'+UsdateTime(StrToDate(DateFin))+'")';
  { FIN PT17 }
  Order := CreateOrderBy;

  StConf := SqlConf ('SALARIES');

  //PGCritere:=Critere;  PT10-2
  SetControlText('STWHERE', Critere);
  if GetControlText('CKRUBNONIMP') = '-' then StRubimpr := 'AND PHB_IMPRIMABLE="X"' else StRubimpr := ''; //PT11
  StRubPat := ' OR PHB_MTPATRONAL<>0 ';                       { PT14 }
  if GetControlText('CKSSRUBPAT') = 'X' then StRubPat :='';   { PT14 }
  if Recap = False then
  begin
    TFQRS1(Ecran).WhereSQL := 'SELECT DISTINCT PHB_RUBRIQUE,PHB_ORDREETAT,PHB_NATURERUB, ' + //PT3-1 Suppression du champ PHB_LIBELLE
    'PSA_SALARIE,PSA_LIBELLE,PHB_SALARIE,PSA_NUMEROSS,PSA_PRENOM,PSA_ADRESSE1,' + AddChamp + ' ' +
      'PSA_ADRESSE3,PSA_ADRESSE2,PSA_CODEPOSTAL,PSA_VILLE,PSA_DATEENTREE,PSA_DATESORTIE, ' +
      'PSA_QUALIFICATION,PSA_COEFFICIENT,PSA_CODEEMPLOI,PSA_PGMODEREGLE,PSA_LIBELLEEMPLOI, ' + {PT1}
    'PCT_LIBELLE,PRM_LIBELLE ' +
      'FROM HISTOBULLETIN ' +
      'LEFT JOIN SALARIES ON PSA_SALARIE=PHB_SALARIE ' +
      'LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ' +
      'LEFT JOIN REMUNERATION ON PHB_NATURERUB=PRM_NATURERUB AND ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE ' +
      'WHERE '+StConf + ' AND (((PHB_MTREM<>0 OR PHB_MTSALARIAL<>0 '+StRubPat+') ' + StRubimpr + ') AND PHB_NATURERUB<>"BAS") ' + //PT11 PT12 { PT14 }
    //'WHERE ((PHB_MTREM+PHB_MTSALARIAL+PHB_MTPATRONAL<>0 AND PHB_IMPRIMABLE="X") AND PHB_NATURERUB<>"BAS") '+//PT2 Modification de la condition
    'AND PHB_DATEFIN>="' + UsDateTime(StrToDate(DateDeb)) + '"  ' + ////PT7 Rmplct DATEDEBUT en DATEFIN
    'AND PHB_DATEFIN<="' + UsDateTime(StrToDate(DateFin)) + '" ' +
      '' + critere + ' ' +
      'ORDER BY ' + Order + '';
  end;

  if Recap = True then
  begin
    StConf := SqlConf ('HISTOBULLETIN');                        // PT13
    TFQRS1(Ecran).WhereSQL := 'SELECT DISTINCT PHB_RUBRIQUE,PHB_ORDREETAT,PHB_NATURERUB, ' +
      '' + AddChamp + 'PCT_LIBELLE,PRM_LIBELLE FROM HISTOBULLETIN ' +
      'LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ' +
      'LEFT JOIN REMUNERATION ON PHB_NATURERUB=PRM_NATURERUB AND ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE ' +
      'WHERE '+StConf + ' AND (((PHB_MTREM<>0 OR PHB_MTSALARIAL<>0 '+StRubPat+') ' + StRubimpr + ') ' + //PT11 { PT14 }
    //  'WHERE ((PHB_MTREM+PHB_MTSALARIAL+PHB_MTPATRONAL<>0 AND PHB_IMPRIMABLE="X") '+ //PT2 Modification de la condition
    'AND (PHB_NATURERUB<>"BAS" )) ' +
      'AND PHB_DATEFIN>="' + UsDateTime(StrToDate(DateDeb)) + '"  ' + //PT7 Rmplct DATEDEBUT en DATEFIN
    'AND PHB_DATEFIN<="' + UsDateTime(StrToDate(DateFin)) + '" ' +
      '' + critere + ' ' +
      'ORDER BY  ' + Order + '';
  end;
  //PgLanceFicInd:=True;  PT10-2
end;

{ PT10-3 Mise en commentaire
procedure TOF_PGFICHEIND.MonnaieInverse(Sender: TObject);
begin
PgMonnaieInv:=(GetControltext('CHMONNAIEINV')='X');
PgTauxConvert:=RendTauxConvertion;
end;}

procedure TOF_PGFICHEIND.ChangeLieuTravail(Sender: TObject);
begin
  RecupChampRupture(Ecran);
  BloqueChampLibre(Ecran);
end;

procedure TOF_PGFICHEIND.ControlPeriodeDeb(Sender: TObject);
begin
  PgIFValidPeriode;
  { PT10-1 Mise en commentaire
  Q := OpenSql('SELECT PEX_DATEDEBUT FROM EXERSOCIAL WHERE PEX_EXERCICE="'+GetControlText('EDEXERSOC')+'" ', True);
  if not Q.eof then DebExer:=Q.findfield('PEX_DATEDEBUT').AsDateTime
  else DebExer:=idate1900;
  Ferme(Q);
  Edit:= ThEdit(getcontrol('XX_VARIABLEDEB'));
  if StrToDate(Edit.Text)<(DebExer) then
     Begin
     PgiBox('La date de début ne peut être inferieur au '+DateToStr(DebExer)+'','Date Erronée!');
     Edit.text:=DateToStr(DebExer);
     End;}
end;

procedure TOF_PGFICHEIND.ControlPeriodeFin(Sender: TObject);
begin
  PgIFValidPeriode;
  { PT10-1 Mise en commentaire
  Q := OpenSql('SELECT PEX_DATEFIN FROM EXERSOCIAL WHERE PEX_EXERCICE="'+GetControlText('EDEXERSOC')+'" ', True);
  if not Q.EOF then // PORTAGECWAS
    FinExer:=Q.findfield('PEX_DATEFIN').AsDateTime
  else
    FinExer:=idate1900;
  Ferme(Q);
  if (StrToDate(GetControlText('XX_VARIABLEFIN'))>FinExer) and (FinExer<>IDate1900) then
     Begin
     PgiBox('La date de fin ne peut être supérieur au '+DateToStr(FinExer)+'','Date Erronée!');
     SetControlText('XX_VARIABLEFIN',DateToStr(FinExer))
     End;         }
end;


procedure TOF_PGFICHEIND.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if edit.text <> '' then
      if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit, 10);
end;
{ DEB PT10-1 }
procedure TOF_PGFICHEIND.PgIFValidPeriode;
var
  YYD, MMD, JJ, YYF, MMF: WORD;
begin
  if IsValidDate(GetControlText('XX_VARIABLEDEB')) and IsValidDate(GetControlText('XX_VARIABLEFIN')) then
  begin
    DecodeDate(StrToDate(GetControlText('XX_VARIABLEDEB')), YYD, MMD, JJ);
    DecodeDate(StrToDate(GetControlText('XX_VARIABLEFIN')), YYF, MMF, JJ);
    if (YYF > YYD) and (MMF >= MMD) then
    begin
      PgiBox('La période d''édition ne peut excéder douze mois civils.', 'Date Erronée!'); { PT10-1 04/05/2004 }
      SetControlText('XX_VARIABLEFIN', DateToStr(FinDeMois(PlusDate(StrToDate(GetControlText('XX_VARIABLEDEB')), 11, 'M'))));
    end;
  end;
end;
{ FIN PT10-1 }

initialization
  registerclasses([TOF_PGFICHEIND]);
end.
