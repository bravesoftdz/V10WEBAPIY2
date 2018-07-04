{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 18/12/2001
Modifié le ... :   /  /
Description .. : Edition des charges sociales
Mots clefs ... : PAIE;CHARGE SOCIALE
*****************************************************************
PT1-1 18/12/2001 SB V571 Fiche de bug n°416 : Edition sous format liste d'exportation
PT1-2 16/01/2002 SB V571 Fiche de bug n°416 : Paie à cheval non éditée
PT2   30/01/2002 SB V571 Edition nominative
PT3   18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non renseigné en Mono
PT4   02/10/2002 SB V585 FQ n°10212 Critère non récupéré pour calcul H/F charges sociales
PT5   20/11/2002 SB V591 Traitements Etats chainés
PT6   18/07/2003 SB V_42 Ajustement edition pour rubrique Loi Fillon
PT7   02/10/2003 SB V_42 Affichage des ongles si gestion paramsoc des combos libres
PT8-1 04/12/2003 SB V_50 FQ 10244 Ajout Rupture périodique
PT8-2 04/12/2003 SB V_50 FQ 10982 Suppression option monnaie inversée
PT9   11/02/2005 SB V_60 FQ 11923 Mode etats chaines, réaffecter l'etat selectionné
PT10  04/07/2005 SB V_60 FQ 12290 Ajout coche pour rub. de régularisations
PT11  22/09/2005 SB V_65 FQ 12277 Ajout coche pour voir cot. sans organisme
PT12  09/05/2007 FC V_72 FQ 13021 Récupération du libellé rubrique dans histobulletin pour avoir celui personnalisé
PT15  06/08/2008 FC V_810 FQ 15562 En CWAS, manque le titre des colonnes dans la liste d'exportation
PT16  22/09/2008 JS FQ 14319 En "état nominatif" possibilité : de trier par nom du salarié
                                                               de sélectionner des matricules salarié
}
unit UTOFPGEditDeclaration;

interface
uses StdCtrls, Classes,sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  eQRS1,
  {$ELSE}
  QRS1,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF,
  ParamDat, ParamSoc, HQry, UTOFPGEtats;

type
  TOF_PGEDIT_CHSOC = class(TOF_PGEtats)
  private
    AjoutChamp, Origine: string;
    DD, DF: TDateTime;
    //procedure MonnaieInverse(Sender: TObject); PT8-2
    procedure Change(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure ExportFiche(Sender: TObject);
    procedure OnClickCkRegul(Sender: TObject); { PT10 }
    procedure ExitEdit(Sender: TObject); //PT16
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure TRI_LIB; //PT16 
  end;
implementation

uses PgEditOutils,PgEditOutils2, EntPaie, PgOutils2; 

{-------------------------------------------------------------------------------
                     TOF  ETAT DES CHARGES SOCIALES
--------------------------------------------------------------------------------}
procedure TOF_PGEDIT_CHSOC.OnArgument(Arguments: string);
var
  Check: TCheckBox;
  DateDeb, DateFin, Defaut, Convert: THEdit;
  Ok: Boolean;
  Min, Max, DebPer, FinPer, ExerPerEncours: string;
  Ckeuro: TCheckBox;
begin
  inherited;
  TFQRS1(Ecran).FCodeEtat := 'PSO';
  //DEB PT5
  DD := idate1900;
  DF := idate1900;
  if (Arguments <> '') then
  begin
    if (Pos('CHAINES', Arguments) > 0) then
    begin
      Origine := ReadTokenSt(Arguments);
      if trim(Arguments) <> '' then
      begin
        DD := StrToDate(ReadTokenSt(Arguments));
        DF := StrToDate(ReadTokenSt(Arguments));
      end;
    end;
  end;
  //FIN PT5

  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    //  Defaut.text:=V_PGI_env.LibDossier;  PT3 Mise en commentaire
    Defaut.text := GetParamSoc('SO_LIBELLE');
  Ckeuro := TCheckBox(GetControl('CKEURO'));
  if Ckeuro <> nil then Ckeuro.Checked := VH_Paie.PGTenueEuro;
  VisibiliteChamp(Ecran); // Compléments visible en fontion des champs SOC organisation
  VisibiliteChampLibre(Ecran);
  { DEB PT7 }
  SetControlProperty('TBCOMPLEMENT', 'Tabvisible', (VH_Paie.PGNbreStatOrg > 0) or (VH_Paie.PGLibCodeStat <> ''));
  SetControlProperty('TBCHAMPLIBRE', 'Tabvisible', (VH_Paie.PgNbCombo > 0));
  { FIN PT7 }
     //Affectation des valeurs par défaut
  DateDeb := ThEdit(getcontrol('XX_VARIABLEDEB'));
  DateFin := ThEdit(getcontrol('XX_VARIABLEFIN'));
  if (DateDeb <> nil) and (DateFin <> nil) then
  begin
    DateDeb.OnElipsisClick := DateElipsisclick;
    DateFin.OnElipsisClick := DateElipsisclick;
  end;
  ok := RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer);
  if Ok = True then
  begin
    if DateDeb <> nil then DateDeb.text := DebPer;
    if DateFin <> nil then DateFin.text := FinPer;
  end;

  RecupMinMaxTablette('CC', 'PGTYPEORGANISME', 'PTG', Min, Max);
  Defaut := ThEdit(getcontrol('ORGANISME'));
  if Defaut <> nil then Defaut.text := Min;
  Defaut := ThEdit(getcontrol('ORGANISME_'));
  if Defaut <> nil then Defaut.text := Max;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Defaut := ThEdit(getcontrol('PHB_ETABLISSEMENT'));
  if Defaut <> nil then Defaut.text := Min;
  Defaut := ThEdit(getcontrol('PHB_ETABLISSEMENT_'));
  if Defaut <> nil then Defaut.text := Max;
  Convert := ThEdit(getcontrol('XX_VARIABLECONV'));
  if convert <> nil then Convert.text := '1';

  // Evenement ONCHANGE
  Check := TCheckBox(GetControl('CN1'));
  if Check <> nil then Check.OnClick := Change;
  Check := TCheckBox(GetControl('CN2'));
  if Check <> nil then Check.OnClick := Change;
  Check := TCheckBox(GetControl('CN3'));
  if Check <> nil then Check.OnClick := Change;
  Check := TCheckBox(GetControl('CN4'));
  if Check <> nil then Check.OnClick := Change;
  Check := TCheckBox(GetControl('CN5'));
  if Check <> nil then Check.OnClick := Change;
  Check := TCheckBox(GetControl('CETAB'));
  if Check <> nil then Check.OnClick := Change;
  Check := TCheckBox(GetControl('CL1'));
  if Check <> nil then Check.OnClick := Change;
  Check := TCheckBox(GetControl('CL2'));
  if Check <> nil then Check.OnClick := Change;
  Check := TCheckBox(GetControl('CL3'));
  if Check <> nil then Check.OnClick := Change;
  Check := TCheckBox(GetControl('CL4'));
  if Check <> nil then Check.OnClick := Change;

  {PT8-2 Mise en commentaire
  ChMonInv := TCheckBox(GetControl('CHMONNAIEINV'));
  if ChMonInv<>nil then ChMonInv.OnClick:=MonnaieInverse;}
  {$IFNDEF EAGLCLIENT}
  TFQRS1(Ecran).FListe.OnClick := ExportFiche; //PT1-1
  {$ENDIF}
  Check := TCheckBox(GetControl('CKNOMI'));
  if Check <> nil then Check.OnClick := ExportFiche;

    //debut PT16
  Check := TCheckBox(GetControl('CKNOMI'));
  if Check <> nil then
   Check.OnClick := Change;
    If TCheckBox(GetControl('CKNOMI')).checked = false then
    begin
      TRI_LIB;
    end;
  Check := TCheckBox(GetControl('CALPHA'));
  if Check <> nil then
   Check.OnClick := Change;
   //fin PT16
  //DEB PT5 Affect critère standard
  if (origine = 'CHAINES') and (DF <> idate1900) then
  begin
    if DateDeb <> nil then DateDeb.text := DateToStr(DD);
    if DateFin <> nil then DateFin.text := DateToStr(DF);
  end;
  //FIN PT5
  { DEB PT8-1 }
  Check := TCheckBox(GetControl('CKPERIODE'));
  if Check <> nil then Check.OnClick := Change;
  SetControlText('XX_RUPTURE2', DateToStr(idate1900));
  { FIN PT8-1 }

   Check := TCheckBox(GetControl('CKREGUL'));                 { PT10 }
   if Assigned(Check) then  Check.OnClick := OnClickCkRegul;  { PT10 }

end;

procedure TOF_PGEDIT_CHSOC.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGEDIT_CHSOC.Change(Sender: TObject);
var
  CEtab, CN1, CN2, CN3, CN4, CN5, CL1, CL2, CL3, CL4: TCheckBox;
  CKNOMI : TCheckBox; //PT16
begin
  BloqueChampLibre(Ecran);
  RecupChampRupture(Ecran);
  AjoutChamp := '';
  SetControltext('XX_RUPTURE2', DateToStr(IDate1900)); // PT8-1
  //Gestion de l'édition charges sociales selon rupture
  CEtab := TCheckBox(GetControl('CETAB'));
  if (CEtab <> nil) then
    if (CEtab.Checked = True) then AjoutChamp := 'PHB_ETABLISSEMENT, ';
  CN1 := TCheckBox(GetControl('CN1'));
  CN2 := TCheckBox(GetControl('CN2'));
  CN3 := TCheckBox(GetControl('CN3'));
  CN4 := TCheckBox(GetControl('CN4'));
  CN5 := TCheckBox(GetControl('CN5'));
  if (CN1 <> nil) and (CN2 <> nil) and (CN3 <> nil) and (CN4 <> nil) and (CN5 <> nil) then
  begin
    if (CN1.Checked = True) then
    begin
      AjoutChamp := 'PHB_TRAVAILN1, ';
    end;
    if (CN2.Checked = True) then
    begin
      AjoutChamp := 'PHB_TRAVAILN2, ';
    end;
    if (CN3.Checked = True) then
    begin
      AjoutChamp := 'PHB_TRAVAILN3, ';
    end;
    if (CN4.Checked = True) then
    begin
      AjoutChamp := 'PHB_TRAVAILN4, ';
    end;
    if (CN5.Checked = True) then
    begin
      AjoutChamp := 'PHB_CODESTAT, ';
    end;
  end;
  //debut PT16
 CKNOMI := TCheckBox(GetControl('CKNOMI'));
 if (CKNOMI <> nil) then
 begin
   If TCheckBox(GetControl('CKNOMI')).checked = true then
   begin
     ThEdit(getcontrol('TPHB_SALARIE')).Visible := true;
     ThEdit(getcontrol('PHB_SALARIE')).Visible := true;
     ThEdit(getcontrol('TPHB_SALARIE_')).Visible := true;
     ThEdit(getcontrol('PHB_SALARIE_')).Visible := true;
     TCheckBox(GetControl('CALPHA')).Visible := true;
   end
   else
   begin
     TRI_LIB;
   end;
 end;
  //fin PT16
  CL1 := TCheckBox(GetControl('CL1'));
  CL2 := TCheckBox(GetControl('CL2'));
  CL3 := TCheckBox(GetControl('CL3'));
  CL4 := TCheckBox(GetControl('CL4'));
  if (CL1 <> nil) and (CL2 <> nil) and (CL3 <> nil) and (CL4 <> nil) then
  begin
    if (CL1.Checked = True) then AjoutChamp := 'PHB_LIBREPCMB1, ';
    if (CL2.Checked = True) then AjoutChamp := 'PHB_LIBREPCMB2, ';
    if (CL3.Checked = True) then AjoutChamp := 'PHB_LIBREPCMB3, ';
    if (CL4.Checked = True) then AjoutChamp := 'PHB_LIBREPCMB4, ';
  end;
  { DEB PT8-1 }
  if GetControlText('CKPERIODE') = 'X' then
  begin
    AjoutChamp := AjoutChamp + 'PHB_DATEFIN, ';
    SetControltext('XX_RUPTURE2', 'PHB_DATEFIN');
  end;
  { FIN PT8-1 }
end;


procedure TOF_PGEDIT_CHSOC.OnUpdate;
var
  SQL, Temp, Tempo, Critere, Org, Orgbis: string;
  Pages: TPageControl;
  DateDeb, DateFin: THEdit;
  x: integer;
  Regul : TCheckBox;
begin
  //-------------------------
  ExportFiche(nil); { PT9 }

  DateDeb := THEdit(GetControl('XX_VARIABLEDEB'));
  DateFin := THEdit(GetControl('XX_VARIABLEFIN'));
  Pages := TPageControl(GetControl('Pages'));
  Temp := RecupWhereCritere(Pages);
  tempo := '';
  critere := '';
  x := Pos('(', Temp);
  if x > 0 then Tempo := copy(Temp, x, (Length(temp) - 5));
  if tempo <> '' then critere := 'AND ' + Tempo;
  { DEB PT10 }
  Regul := TCheckBox(GetControl('CKREGUL'));
  if Assigned(Regul) then
   IF Regul.State = cbUnchecked then
     Critere := Critere + ' AND PHB_COTREGUL<>"REG" '
  Else
   IF Regul.State = cbChecked then
     Critere := Critere + ' AND PHB_COTREGUL="REG" ';
  { FIN PT10 }
  { DEB PT11 }
  Org := GetControlText('ORGANISME');
  Orgbis:= GetControlText('ORGANISME_');
  if (GetControlText('CKCOTSSORGANISME')='X') then
     Begin
     if ( Org <> '') AND (Orgbis <> '') then
       Critere := Critere + ' AND ((PHB_ORGANISME>="'+Org+'" AND PHB_ORGANISME <="'+OrgBis+'") OR PHB_ORGANISME="") '
     else
       if ( Org = '') AND (Orgbis <> '') then
       Critere := Critere + ' AND (PHB_ORGANISME <="'+OrgBis+'" OR PHB_ORGANISME="") '
     else
       if ( Org <> '') AND (Orgbis = '') then
      Critere := Critere + ' AND (PHB_ORGANISME>="'+Org+'" OR PHB_ORGANISME="") '
     End
  else
    if (GetControlText('CKCOTSSORGANISME')='-') then
      Begin
      if ( Org <> '') then Critere := Critere + ' AND PHB_ORGANISME>="'+Org+'" ';
      if ( Orgbis <> '') then Critere := Critere + ' AND PHB_ORGANISME<="'+Orgbis+'" ';
      end;
  { FIN PT11 }


  SetControlText('CRITERE', Critere); //PT4 Affect critere pour recup Fn



  if (IsValidDate(DateDeb.text)) and (IsValidDate(DateFin.text)) then
  begin
    if GetControlText('CKNOMI') = '-' then
    begin
      SQL := 'SELECT PHB_RUBRIQUE,PHB_NATURERUB,PHB_ORGANISME,PHB_COTREGUL,PHB_LIBELLE as PCT_LIBELLE,PCT_PREDEFINI,PCT_REDUCBASSAL,PHB_TAUXSALARIAL,' + AjoutChamp + //PT6 PT10 18/08/2005 //PT12
      ' PHB_TAUXPATRONAL,COUNT(DISTINCT PHB_SALARIE) AS NBSAL,SUM(PHB_BASECOT) AS BASECOT,' +
        ' SUM(PHB_MTSALARIAL)AS MTSAL,SUM(PHB_MTPATRONAL)AS MTPAT FROM HISTOBULLETIN ' +
        ' LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ' +
        ' WHERE PHB_NATURERUB="COT" AND (PHB_MTSALARIAL<>0 OR PHB_MTPATRONAL<>0)  ' +
        ' AND PHB_DATEFIN>="' + UsDateTime(StrToDate(DateDeb.text)) + '" ' + //PT1-2 Chgmt PHB_DATEDEBUT
      ' AND PHB_DATEFIN<="' + UsDateTime(StrToDate(DateFin.text)) + '" ' +
        ' ' + Critere + ' ' +
        ' GROUP BY ' + AjoutChamp + 'PHB_ORGANISME,PHB_NATURERUB,PHB_RUBRIQUE,PHB_COTREGUL,PCT_PREDEFINI,PHB_LIBELLE,PCT_REDUCBASSAL,PHB_TAUXSALARIAL,PHB_TAUXPATRONAL' + //PT6 PT10 18/08/2005    //PT12
      ' ORDER BY ' + AjoutChamp + 'PHB_ORGANISME,PHB_RUBRIQUE,PHB_LIBELLE ';  //PT12
    end
    else
    begin
    //debut PT16
     If TCheckBox(GetControl('CALPHA')).Checked = True then
      begin
      SQL := 'SELECT PSA_LIBELLE, PHB_SALARIE,PHB_RUBRIQUE,PHB_NATURERUB,PHB_ORGANISME,PHB_COTREGUL,PCT_PREDEFINI,' +
        'PHB_LIBELLE as PCT_LIBELLE,PHB_TAUXSALARIAL,PHB_TAUXPATRONAL,' + AjoutChamp +
        'SUM(PHB_BASECOT) AS BASECOT,SUM(PHB_MTSALARIAL)AS MTSAL,SUM(PHB_MTPATRONAL)AS MTPAT ' +
        'FROM HISTOBULLETIN ' +
        'LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ' +
        'LEFT JOIN SALARIES ON PHB_SALARIE=PSA_SALARIE '+
        ' WHERE PHB_NATURERUB="COT" AND (PHB_MTSALARIAL<>0 OR PHB_MTPATRONAL<>0)  ' +
        ' AND PHB_DATEFIN>="' + UsDateTime(StrToDate(DateDeb.text)) + '" ' +
        ' AND PHB_DATEFIN<="' + UsDateTime(StrToDate(DateFin.text)) + '" ' +
        ' ' + Critere + ' ' +
        'GROUP BY ' + AjoutChamp + 'PHB_ORGANISME,PHB_SALARIE,PHB_NATURERUB,PHB_RUBRIQUE,'+
        'PSA_libelle, PHB_COTREGUL,PCT_PREDEFINI,PHB_LIBELLE,PHB_TAUXSALARIAL,PHB_TAUXPATRONAL ' +
        'ORDER BY ' + AjoutChamp + 'PHB_ORGANISME,PHB_RUBRIQUE,PSA_libelle, PHB_LIBELLE,PHB_SALARIE';
      end
     else //fin PT16
      begin
      SQL := 'SELECT PHB_SALARIE,PHB_RUBRIQUE,PHB_NATURERUB,PHB_ORGANISME,PHB_COTREGUL,PCT_PREDEFINI,' + //PT10 18/08/2005
        'PHB_LIBELLE as PCT_LIBELLE,PHB_TAUXSALARIAL,PHB_TAUXPATRONAL,' + AjoutChamp +   //PT12
        'SUM(PHB_BASECOT) AS BASECOT,SUM(PHB_MTSALARIAL)AS MTSAL,SUM(PHB_MTPATRONAL)AS MTPAT ' +
        'FROM HISTOBULLETIN ' +
        'LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ' +
        ' WHERE PHB_NATURERUB="COT" AND (PHB_MTSALARIAL<>0 OR PHB_MTPATRONAL<>0)  ' +
        ' AND PHB_DATEFIN>="' + UsDateTime(StrToDate(DateDeb.text)) + '" ' + //PT1-2 Chgmt PHB_DATEDEBUT
      ' AND PHB_DATEFIN<="' + UsDateTime(StrToDate(DateFin.text)) + '" ' +
        ' ' + Critere + ' ' +
        'GROUP BY ' + AjoutChamp + 'PHB_ORGANISME,PHB_SALARIE,PHB_NATURERUB,PHB_RUBRIQUE,'+
        'PHB_COTREGUL,PCT_PREDEFINI,PHB_LIBELLE,PHB_TAUXSALARIAL,PHB_TAUXPATRONAL ' + //PT10 18/08/2005    //PT12
        'ORDER BY ' + AjoutChamp + 'PHB_ORGANISME,PHB_RUBRIQUE,PHB_LIBELLE,PHB_SALARIE'; //PT12
      end;
    end;
    TFQRS1(Ecran).WhereSQL := SQL;
  end
  else
    HShowMessage('5;Saisie :;Vous devez saisir une période de début et de fin de paie!;W;O;O;O;;;', '', '');
end;

{PT8-2 Mise en commentaire
procedure TOF_PGEDIT_CHSOC.MonnaieInverse(Sender: TObject);
var
ChMonInv : TCheckBox;
Convert:ThEdit;
begin
ChMonInv := TCheckBox(GetControl('CHMONNAIEINV'));
Convert:=ThEdit(getcontrol('XX_VARIABLECONV'));
if (ChMonInv<>nil) and (convert<>nil)  then
  begin
  if ChMonInv.checked=False then Convert.text:='1';
  if ChMonInv.checked=True  then Convert.text:=RendTauxConvertion;
  End;
end;}

//DEB PT1-1
procedure TOF_PGEDIT_CHSOC.ExportFiche(Sender: TObject);
var
  s: string;
begin
//PT15  {$IFNDEF EAGLCLIENT}
  if TFQRS1(Ecran).FListe.Checked then
  begin
    if GetControlText('CKNOMI') = '-' then
      s := 'PSE'
    else
      s := 'PNE';
  end
  else
//PT15    {$ENDIF}
  begin
    if GetControlText('CKNOMI') = '-' then
      s := 'PSO'
    else
      s := 'PSN';
  end;
  TFQRS1(Ecran).FCodeEtat := s;
end;
//FIN PT1-1
procedure TOF_PGEDIT_CHSOC.OnLoad;
begin
  inherited;
  //DEB PT5 Affect critère standard
  if (origine = 'CHAINES') and (DF <> idate1900) then
  begin
    SetControlText('XX_VARIABLEDEB', DateToStr(DD));
    SetControlText('XX_VARIABLEFIN', DateToStr(DF));
  end;
  //FIN PT5
end;

{ DEB PT10 }
procedure TOF_PGEDIT_CHSOC.OnClickCkRegul(Sender: TObject);
begin
  If Sender is TCheckBox then
    IF TCheckBox(Sender).State = cbUnchecked then
       TCheckBox(Sender).Caption := 'Sans rubriques de régularisations'
    Else
      IF TCheckBox(Sender).State = cbChecked then
        TCheckBox(Sender).Caption := 'Uniquement rubriques de régularisations (*R)'
      Else
        IF TCheckBox(Sender).State = cbGrayed then
          TCheckBox(Sender).Caption := 'Avec rubriques de régularisations (*R)';
end;
{ FIN PT10 }

//debut PT16
procedure TOF_PGEDIT_CHSOC.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if edit.text <> '' then
      if (VH_Paie.PgTypeNumSal = 'NUM') and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PGEDIT_CHSOC.TRI_LIB();
var
 Min, Max : string;
 Defaut : THEdit;
begin
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
 ThEdit(getcontrol('TPHB_SALARIE')).Visible := false;
 ThEdit(getcontrol('PHB_SALARIE')).Visible := false;
 ThEdit(getcontrol('TPHB_SALARIE_')).Visible := false;
 ThEdit(getcontrol('PHB_SALARIE_')).Visible := false;
 TCheckBox(GetControl('CALPHA')).Visible := false;
 TCheckBox(GetControl('CALPHA')).Checked := false;
end;
//fin PT16

initialization
  registerclasses([TOF_PGEDIT_CHSOC]);
end.

