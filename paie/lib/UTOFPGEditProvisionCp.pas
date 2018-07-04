{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : Provision & Solde CP
Mots clefs ... : PAIE;CP
*****************************************************************
PT1 11/09/2001 V547 SB Gestion de la monnaie de tenue
PT2 28/01/2002 V571 SB Sélection alphanumérique du code salarié
PT3 29/01/2002 V571 SB Idate1900 au lieu de 01/01/1900
PT4 01/03/2002 V571 SB Appel de la Moulinette des maj des bases des mvts d'acquis congés payés
PT5 18/04/2002 V571 SB Fiche de bug n°10087 : V_PGI_env.LibDossier non renseigné en Mono
PT6 18/12/2002 V585 SB FQ 10409 Seuls les critères salarié sont pris en compte au chargement des enr. absences
PT7 13/06/2003 V_42 SB On exclut les salariés pas encore entrée..
PT8 02/10/2003 V_42 SB Affichage des ongles si gestion paramsoc des combos libres
PT9 02/04/2004 V_50 PH Prise en compte de la provision CP et RTT faite à partir de la table PROVCP
PT10 09/04/2004 V_50 SB FQ 11136 Ajout Gestion des congés payés niveau salarié
PT11 31/08/2004 V_50 PH FQ 11556 ergonomie titre de la fénêtre
PT12 22/04/2005 V_60 SB FQ 12207 Solde CP : Refonte tri alpha après tri etablissement
PT13 29/04/2005 V_60 PH FQ 12224 Edition CP et RTT Menu dotations CP et RTT
PT14 22/06/2005 V_60 PH FQ 11648 Initialisation de la date d'arrêté à la date de fin de mois
PT15 24/04/2006 V_65 SB FQ 12406 Ajout niveau de rupture
PT16 20/07/2006 V_65 SB FQ 13325 Refonte Tri
PT17 21/03/2007 V_70 FC FQ 13981 Prendre en compte les salariés avec une date de sortie 30/12/1899
PT18 25/06/2007 V_70 FC FQ 11556 Ergonomie

}
unit UTOFPGEditProvisionCp;

interface
uses StdCtrls, Controls, Classes,  sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  eQRS1,
  {$ELSE}
  QRS1,
  {$ENDIF}
  HCtrls, HEnt1, UTOF, ParamDat, //HMsgBox{$IFNDEF DBXPRESS} ,dbTables {$ELSE} ,uDbxDataSet {$ENDIF} forms,  Graphics,
  ParamSoc, HQry;

type
  TOF_PGPROVISIONCP = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
  private
    TypeProv: string; // PT9
    procedure DateElipsisclick(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    procedure ChangeAlpha(Sender: TObject);
    procedure OnChangeRupture(Sender: TObject);   { PT15 }
  end;

  TOF_PGSOLDECP_ETAT = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
  private
    AjoutChamp, AjoutTri: string;
    procedure DateElipsisclick(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    procedure ChangeLieuTravail(Sender: TObject);
  end;

implementation

uses PgEditOutils, PGEditOutils2,EntPaie, PGoutils2;


{ TOF_PGPROVISIONCP }

procedure TOF_PGPROVISIONCP.OnArgument(Arguments: string);
var
  Edit: THEdit;
  Ckeuro: TCheckBox;
  Min, Max: string;
  Check: TCheckBox;
  Combo : ThValComboBox;
  i     : integer;
begin
  inherited;
  //PT9 02/04/2004 V_50 PH Prise en compte de la provision CP et RTT faite à partir de la table PROVCP
  TypeProv := readTokenSt(Arguments);
  if TypeProv = 'PRD' then Ecran.Caption := 'Récapitulatif CP et RTT'; // PT11
  UpdateCaption (Ecran); // PT18
  VisibiliteChamp(Ecran);
  VisibiliteChampLibre(Ecran);
  { DEB PT8 }
  SetControlProperty('TBCOMPLEMENT', 'Tabvisible', (VH_Paie.PGNbreStatOrg > 0) or (VH_Paie.PGLibCodeStat <> ''));
  SetControlProperty('TBCHAMPLIBRE', 'Tabvisible', (VH_Paie.PgNbCombo > 0));
  { FIN PT8 }
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Edit := ThEdit(getcontrol('PSA_SALARIE'));
  if Edit <> nil then
  begin
    Edit.text := Min;
    Edit.OnExit := ExitEdit;
  end;
  Edit := ThEdit(getcontrol('PSA_SALARIE_'));
  if Edit <> nil then
  begin
    Edit.text := Max;
    Edit.OnExit := ExitEdit;
  end;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Edit := ThEdit(getcontrol('PSA_ETABLISSEMENT'));
  if Edit <> nil then Edit.text := Min;
  Edit := ThEdit(getcontrol('PSA_ETABLISSEMENT_'));
  if Edit <> nil then Edit.text := Max;
  // DEB PT14
  if TypeProv = '' then Edit := THEdit(GetControl('DATEARRET'))
  else Edit := THEdit(GetControl('_DATEFIN'));
  // FIN PT14
  if Edit <> nil then
  begin
    Edit.OnDblClick := DateElipsisclick;
    if TypeProv = '' then Edit.Text := DateToStr(Date); // PT14
  end;
  Edit := ThEdit(getcontrol('DOSSIER'));
  if Edit <> nil then
    //  Edit.text:=V_PGI_env.LibDossier; //PT5 Mise en commentaire
    Edit.text := GetParamSoc('SO_LIBELLE');
  Ckeuro := TCheckBox(GetControl('CKEURO')); //PT1
  if Ckeuro <> nil then Ckeuro.Checked := VH_Paie.PGTenueEuro;
  Check := TCheckBox(GetControl('CALPHA')); //PT2
  if Check <> nil then Check.OnClick := ChangeAlpha;
  { DEB PT15 }
  Combo := ThValComboBox(GetControl('THVALRUPTURE1'));
  if Assigned(Combo) then
     Begin
     Combo.onchange := OnChangeRupture;
     Combo.Itemindex := 0;
     Combo.Items.Add('<<Aucune>>');          Combo.Values.Add('');
     For i :=1 To VH_Paie.PGNbreStatOrg Do
       Begin
       IF (i= 1) AND (VH_Paie.PGLibelleOrgStat1 <> '') then Combo.Items.Add(VH_Paie.PGLibelleOrgStat1)
       else IF (i= 2) AND (VH_Paie.PGLibelleOrgStat2 <> '') then Combo.Items.Add(VH_Paie.PGLibelleOrgStat2)
       else IF (i= 3) AND (VH_Paie.PGLibelleOrgStat3 <> '') then Combo.Items.Add(VH_Paie.PGLibelleOrgStat3)
       else IF (i= 4) AND (VH_Paie.PGLibelleOrgStat4 <> '') then Combo.Items.Add(VH_Paie.PGLibelleOrgStat4);
       Combo.Values.Add('PSA_TRAVAILN'+IntToStr(i));
       End;
     if VH_Paie.PGLibCodeStat <> '' then
       Begin
       Combo.Items.Add(VH_Paie.PGLibCodeStat);
       Combo.Values.Add('PSA_CODESTAT');
       End;
     For i :=1 To VH_Paie.PgNbCombo Do
       Begin
       IF (i= 1) AND (VH_Paie.PgLibCombo1 <> '') then Combo.Items.Add(VH_Paie.PgLibCombo1)
       else IF (i= 2) AND (VH_Paie.PgLibCombo2 <> '') then Combo.Items.Add(VH_Paie.PgLibCombo2)
       else IF (i= 3) AND (VH_Paie.PgLibCombo3 <> '') then Combo.Items.Add(VH_Paie.PgLibCombo3)
       else IF (i= 4) AND (VH_Paie.PgLibCombo4 <> '') then Combo.Items.Add(VH_Paie.PgLibCombo4);
       Combo.Values.Add('PSA_LIBREPCMB'+IntToStr(i));
       End;
     End;
  { FIN PT15 }

end;

procedure TOF_PGPROVISIONCP.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;


procedure TOF_PGPROVISIONCP.OnUpdate;
var
  SQL, Temp, Tempo, Critere, {Condition,PT6} StTri, st, StRupt: string;
  DD: TDateTime;
  Pages: TPageControl;
  x: integer;
begin
  inherited;
  //MAJBaseCPAcquis(True); //PT4  integré avant calcul provision
  if GetControlText('CALPHA') = 'X' then StTri := 'PSA_LIBELLE,' else StTri := '';
  Pages := TPageControl(GetControl('Pages'));
  Temp := RecupWhereCritere(Pages);
  tempo := '';
  critere := '';
  x := Pos('(', Temp);
  if x > 0 then Tempo := copy(Temp, x, (Length(temp) - 5));
  if tempo <> '' then critere := 'AND ' + Tempo;
  SetControlText('STWHERE', critere); //PT6
  { DEB PT15 }
  If GetControlText('XX_RUPTURE1') <> '' then StRupt := GetControlText('XX_RUPTURE1')+','
  else StRupt :='';
  { FIN PT15 }
  // DEB PT14
  if TypeProv = '' then DD := StrToDate(GetControlText('DATEARRET'))
  else DD := StrToDate(GetControlText('DATEFIN'));
  // FIN PT14
  //PT9 02/04/2004 V_50 PH Prise en compte de la provision CP et RTT faite à partir de la table PROVCP
  st := ' ' + critere + ' ' + //PT3
  'AND (PSA_DATESORTIE>="' + USDateTime(DD) + '" OR PSA_DATESORTIE<="' + UsdateTime(Idate1900) + '" OR PSA_DATESORTIE is null) '; //PT17
// PT13 On prend en compte tous les salariés calculés dans PROVCP car on peut calculer avec effet rétroactif-
// Cas entrée/sortie/entrée
  if TypeProv = '' then st := st + 'AND PSA_DATEENTREE<="' + USDateTime(DD) + '" '; //PT7
// FIN PT13
  if (DD > idate1900) then
  begin
    if TypeProv = '' then
      SQL := 'SELECT PSA_SALARIE,PSA_ETABLISSEMENT,'+StRupt+'PSA_LIBELLE,PSA_PRENOM,ETB_LIBELLE ' + { PT15 }
        'FROM SALARIES ' +
        'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
        'WHERE ETB_CONGESPAYES="X" AND PSA_CONGESPAYES="X" ' + st + { PT10 }
      {     // PT9  ' ' + critere + ' ' + //PT3
          'AND (PSA_DATESORTIE>="' + USDateTime(DD) + '" OR PSA_DATESORTIE="' + UsdateTime(Idate1900) + '" OR PSA_DATESORTIE is null) ' +
          'AND PSA_DATEENTREE<="' + USDateTime(DD) + '" ' + //PT7 // PT9 }
      'ORDER BY PSA_ETABLISSEMENT,' + StRupt +  StTri + 'PSA_SALARIE'  { PT15 }
    else
    begin
      SQL := 'SELECT PSA_SALARIE,PSA_ETABLISSEMENT,PSA_LIBELLE,PSA_PRENOM,ETB_LIBELLE,PROVCP.* ' +
        'FROM PROVCP LEFT JOIN SALARIES ON PSA_SALARIE=PDC_SALARIE ' +
        'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
        'WHERE ETB_CONGESPAYES="X" AND PSA_CONGESPAYES="X" ' + st + ' AND PDC_DATEARRET="' + USDateTime(DD) + '" ' + { PT10 }
      'ORDER BY PSA_ETABLISSEMENT,' + StTri + 'PSA_SALARIE';
      TFQRS1(Ecran).CodeEtat := TypeProv;
    end;
  end;
  TFQRS1(Ecran).WhereSQL := SQL;

  {PT6 mise en commentaire
  Condition:='';
  if GetControlText('PSA_SALARIE')<>''  then Condition:=' AND PSA_SALARIE>="'+GetControlText('PSA_SALARIE')+'" ';
  if GetControlText('PSA_SALARIE_')<>'' then Condition:=Condition+' AND PSA_SALARIE<="'+GetControlText('PSA_SALARIE_')+'" ';
  SetControlText('STWHERE',Condition);       }
end;

procedure TOF_PGPROVISIONCP.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PGPROVISIONCP.ChangeAlpha(Sender: TObject);
begin
  if TCheckBox(Sender).Name = 'CALPHA' then //PT2
    AffectCritereAlpha(Ecran, TCheckBox(Sender).Checked, 'PSA_SALARIE', 'PSA_LIBELLE');
end;


{ DEB PT15 }
procedure TOF_PGPROVISIONCP.OnChangeRupture(Sender: TObject);
var
  Combo: THValComboBox;
  ChampRupture, Name, Ordre: string;
  IntOrdre, i: integer;
begin
  if Sender is THValComboBox then Combo := THValComboBox(Sender) else exit;
  Name := Combo.name;
  Ordre := Copy(name, Length(name), 1);
  ChampRupture := 'XX_RUPTURE' + Ordre;
  if Combo.value = '' then
    SetControlText(ChampRupture, '')
  else
    SetControlText(ChampRupture, Combo.value);
end;
{ FIN PT15 }

{ TOF_PGSOLDECP_ETAT }

procedure TOF_PGSOLDECP_ETAT.OnArgument(Arguments: string);
var
  Edit: THEdit;
  Min, Max: string;
  Check: TCheckBox;
begin
  inherited;
  VisibiliteChamp(Ecran);
  VisibiliteChampLibre(Ecran);
  { DEB PT8 }
  SetControlProperty('TBCOMPLEMENT', 'Tabvisible', (VH_Paie.PGNbreStatOrg > 0) or (VH_Paie.PGLibCodeStat <> ''));
  SetControlProperty('TBCHAMPLIBRE', 'Tabvisible', (VH_Paie.PgNbCombo > 0));
  { FIN PT8 }
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Edit := ThEdit(getcontrol('PSA_SALARIE'));
  if Edit <> nil then
  begin
    Edit.text := Min;
    Edit.OnExit := ExitEdit;
  end;
  Edit := ThEdit(getcontrol('PSA_SALARIE_'));
  if Edit <> nil then
  begin
    Edit.text := Max;
    Edit.OnExit := ExitEdit;
  end;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Edit := ThEdit(getcontrol('PSA_ETABLISSEMENT'));
  if Edit <> nil then Edit.text := Min;
  Edit := ThEdit(getcontrol('PSA_ETABLISSEMENT_'));
  if Edit <> nil then Edit.text := Max;
  Edit := THEdit(GetControl('DATEARRET'));
  if Edit <> nil then
  begin
    Edit.OnDblClick := DateElipsisclick;
    Edit.Text := DateToStr(Date);
  end;
  Edit := ThEdit(getcontrol('DOSSIER'));

  if Edit <> nil then
    //  Edit.text:=V_PGI_env.LibDossier;    //PT5 Mise en commentaire
    Edit.text := GetParamSoc('SO_LIBELLE');
  //Evenement ONCHANGE
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
  Check := TCheckBox(GetControl('CALPHA'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL1'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL2'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL3'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL4'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;

end;

procedure TOF_PGSOLDECP_ETAT.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGSOLDECP_ETAT.OnUpdate;
var
  SQL, Temp, Tempo, Critere,AjoutAlpha {PT6,Condition}: string;
  DD: TDateTime;
  Pages: TPageControl;
  x: integer;
begin
  inherited;
  Pages := TPageControl(GetControl('Pages'));
  if Pages <> nil then Temp := RecupWhereCritere(Pages);
  tempo := '';
  critere := '';
  x := Pos('(', Temp);
  if x > 0 then Tempo := copy(Temp, x, (Length(temp) - 5));
  if tempo <> '' then critere := 'AND ' + Tempo;
  if GetControlText('CALPHA')='X' then AjoutAlpha := ',PSA_LIBELLE,PSA_PRENOM';   { PT12 }
  {PT6
  Condition:='';
  if GetControlText('PSA_SALARIE')<>''  then Condition:=' AND PSA_SALARIE>="'+GetControlText('PSA_SALARIE')+'" ';
  if GetControlText('PSA_SALARIE_')<>'' then Condition:=Condition+' AND PSA_SALARIE<="'+GetControlText('PSA_SALARIE_')+'" ';}
  SetControlText('STWHERE', critere);
  DD := StrToDate(GetControlText('DATEARRET'));
  if (DD > idate1900) then
    SQL := 'SELECT PSA_SALARIE,PSA_ETABLISSEMENT,PSA_LIBELLE,PSA_PRENOM,' + AjoutChamp + 'ETB_LIBELLE ' +
      'FROM SALARIES ' +
      'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
      'WHERE ETB_CONGESPAYES="X" AND PSA_CONGESPAYES="X" ' + { PT10 }
    ' ' + critere + ' ' + //PT3
    'AND (PSA_DATESORTIE>="' + USDateTime(DD) + '" OR PSA_DATESORTIE<="' + UsdateTime(Idate1900) + '" OR PSA_DATESORTIE is null) ' + //PT17
      'AND PSA_DATEENTREE<="' + USDateTime(DD) + '" ' + //PT7
    'ORDER BY ' + AjoutTri + 'PSA_ETABLISSEMENT'+AjoutAlpha+',PSA_SALARIE';   { PT12 } { PT16 }
  TFQRS1(Ecran).WhereSQL := SQL;
end;

procedure TOF_PGSOLDECP_ETAT.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PGSOLDECP_ETAT.ChangeLieuTravail(Sender: TObject);
var
  CN1, CN2, CN3, CN4, CN5, CL1, CL2, CL3, CL4 : TCheckBox;
begin
  AjoutChamp := '';
  BloqueChampLibre(Ecran);
  RecupChampRupture(Ecran);
  SetControlText('STCHAMP', GetControlText('XX_RUPTURE1'));
  CN1 := TCheckBox(GetControl('CN1'));
  CN2 := TCheckBox(GetControl('CN2'));
  CN3 := TCheckBox(GetControl('CN3'));
  CN4 := TCheckBox(GetControl('CN4'));
  CN5 := TCheckBox(GetControl('CN5'));
  CL1 := TCheckBox(GetControl('CL1'));
  CL2 := TCheckBox(GetControl('CL2'));
  CL3 := TCheckBox(GetControl('CL3'));
  CL4 := TCheckBox(GetControl('CL4'));
//  Alpha := TCheckBox(GetControl('CALPHA')); PT12

  if (CN1 <> nil) and (CN2 <> nil) and (CN3 <> nil) and (CN4 <> nil) and (CN5 <> nil) then
  begin
    if (CN1.Checked = True) then AjoutChamp := 'PSA_TRAVAILN1,';
    if (CN2.Checked = True) then AjoutChamp := 'PSA_TRAVAILN2,';
    if (CN3.Checked = True) then AjoutChamp := 'PSA_TRAVAILN3,';
    if (CN4.Checked = True) then AjoutChamp := 'PSA_TRAVAILN4,';
    if (CN5.Checked = True) then AjoutChamp := 'PSA_CODESTAT,';
  end;
  if (CL1 <> nil) and (CL2 <> nil) and (CL3 <> nil) and (CL4 <> nil) then
  begin
    if (CL1.Checked = True) then AjoutChamp := 'PSA_LIBREPCMB1,';
    if (CL2.Checked = True) then AjoutChamp := 'PSA_LIBREPCMB2,';
    if (CL3.Checked = True) then AjoutChamp := 'PSA_LIBREPCMB3,';
    if (CL4.Checked = True) then AjoutChamp := 'PSA_LIBREPCMB4,';
  end;

  AjoutTri := AjoutChamp;

 { PT12 Mise en commentaire
 // TRI SUR LIBELLE SALARIE + CHAMPS RUPTURE    //Cas Combiné avec tri alphabétique
  if Alpha <> nil then
    if Alpha.Checked = True then
    begin
      AjoutTri := 'PSA_LIBELLE,';
    end;
  if (CN1 <> nil) and (CN2 <> nil) and (CN3 <> nil) and (CN4 <> nil) and (CN5 <> nil) and (Alpha <> nil) then
  begin
    if (CN1.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PSA_TRAVAILN1,PSA_LIBELLE,';
    if (CN2.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PSA_TRAVAILN2,PSA_LIBELLE,';
    if (CN3.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PSA_TRAVAILN3,PSA_LIBELLE,';
    if (CN4.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PSA_TRAVAILN4,PSA_LIBELLE,';
    if (CN5.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PSA_CODESTAT,PSA_LIBELLE,';
  end;
  if (CL1 <> nil) and (CL2 <> nil) and (CL3 <> nil) and (CL4 <> nil) and (Alpha <> nil) then
  begin
    if (CL1.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PSA_LIBREPCMB1,PSA_LIBELLE,';
    if (CL2.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PSA_LIBREPCMB2,PSA_LIBELLE,';
    if (CL3.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PSA_LIBREPCMB3,PSA_LIBELLE,';
    if (CL4.Checked = True) and (Alpha.Checked = True) then AjoutTri := 'PSA_LIBREPCMB4,PSA_LIBELLE,';
  end;   }

  if TCheckBox(Sender).Name = 'CALPHA' then //PT2
    AffectCritereAlpha(Ecran, TCheckBox(Sender).Checked, 'PSA_SALARIE', 'PSA_LIBELLE');

end;

initialization
  registerclasses([TOF_PGPROVISIONCP, TOF_PGSOLDECP_ETAT]);
end.

