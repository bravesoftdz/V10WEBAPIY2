{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 22/03/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDIT_ENFANTS ()
Mots clefs ... : TOF;PGEDIT_ENFANTS
*****************************************************************
PT1 18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non renseigné en Mono
PT2 07/08/2002 JL V582 Fiche de bug n° 10169 Modification du titre de la fiche
PT3 25/06/2004 JL V_50 Fiche de bug n° 11350 Inclure salarié avec date sortie=30/12/1899
PT4 09/03/2005 JL V_60 FQ 12086 Corrections si tri alphabétique
PT5 09/03/2005 JL V_60 FQ 12395 Ajout choix salarié present
PT6 07/02/2007 FC V_80 Mise en place filtrage habilitations/populations
PT7 26/03/2007 GGS v_80 FQ 14012 Prise en compte sélection de matricules
PT8 21/05/2007 FC V_72 FQ 14251 Prise en compte des critères
PT9 30/05/2007 FC V_72 FQ 12925 Rajout du type de lien de parenté
PT10 15/06/2007 FC V_72 FQ 14244 partir de la table SALARIES et non ENFANTSALARIE pour la gestion de la confidentialité
PT11 20/06/2007 FC V_72 FQ 14415 Modification du titre de la fiche
}
unit UTofPgEdit_Enfants;

interface

uses StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1,
  {$ELSE}
  eQRS1,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, P5Def, EntPaie,
  PGoutils,PGoutils2, PGEditOutils,PgEditOutils2, HQry,
  Paramsoc;

type
  TOF_PGEDIT_ENFANTS = class(TOF)
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
  private
    TriCN: string;
    procedure ExitEdit(Sender: TObject);
    procedure ChangeLieuTrav(Sender: TObject);
    procedure CkSortieClick (Sender : TObject);
  end;

implementation

procedure TOF_PGEDIT_ENFANTS.OnUpdate;
var
  Where, Tri, SQL, TriEtab, VerifSortie: string;
  Pages: TPageControl;
  Habilitation:String;//PT11
begin
  inherited;
  TriEtab := '';
  Tri := '';
  Pages := TPageControl(GetControl('Pages')); //PT8
  Where := RecupWhereCritere(Pages);
  //DEBUT PT5
  If GetCheckBoxState('CKSORTIE') = CbChecked then
  begin
       if Where <> '' then VerifSortie := ' AND (PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(StrToDate(GetControlText('DATESORTIE')))+'")'
       else VerifSortie := ' (PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(StrToDate(GetControlText('DATESORTIE')))+'")'
  end
  else VerifSortie := '';
  // FIN PT5
  if Where <> '' then Where := Where + VerifSortie
  else Where := 'WHERE ' + VerifSortie;
  if TriCN = '' then TriCN := ' ORDER BY ';
  if GetControlText('CE') = 'X' then
  begin
    TriEtab := TriCN + 'PSA_ETABLISSEMENT,';
    SetControlText('XX_RUPTURE3', 'PSA_ETABLISSEMENT');
  end
  else
  begin
    TriEtab := TriCN;
    SetControlText('XX_RUPTURE3', '');
  end;
  if GetControlText('ALPHA') = 'X' then
  begin
    Tri := TriEtab + 'PSA_LIBELLE,PSA_SALARIE';//PT4
    SetControlText('XX_RUPTURE2', 'PSA_SALARIE');
  end
  else
  begin
    Tri := TriEtab + 'PSA_SALARIE';
    SetControlText('XX_RUPTURE2', 'PSA_SALARIE');
  end;
  SQL := 'SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,' +
    ' PEF_NOM,PEF_PRENOM,PEF_NATIONALITE,PEF_SEXE,PEF_DATENAISSANCE,PEF_ACHARGE,PEF_TYPEPARENTAL FROM SALARIES LEFT JOIN' +  //PT9
    ' ENFANTSALARIE ON PEF_SALARIE=PSA_SALARIE '; //PT10
//PT8 Mise en commentaire du PT7 : récupération des critères via RecupWhereCritere qui fonctionnait mal => corrigé
{//PT7
  if (GetControlText('PSA_SALARIE') <> '') and (GetControlText('PSA_SALARIE_') <> GetControlText('PSA_SALARIE')) then
    SQL:=SQL + 'WHERE PEF_SALARIE>= "' + GetControlText('PSA_SALARIE') +'" AND PEF_SALARIE<= "' + GetControlText('PSA_SALARIE_') +'"';
  if (GetControlText('PSA_SALARIE') <> '') and (GetControlText('PSA_SALARIE_') = GetControlText('PSA_SALARIE')) then
    SQL:=SQL + ' WHERE PEF_SALARIE= "' + GetControlText('PSA_SALARIE_') +'"';
//FIN PT7}

  //PT8 Mise en commentaire du PT6 : récupération des critères via RecupWhereCritere qui fonctionnait mal => corrigé pour prendre en compte les habilitations
  {//DEB PT6
  Habilitation := '';
  if Assigned(MonHabilitation) and (MonHabilitation.LeSQL<>'') then
  begin
    if (Where<>'') and (Where<>'WHERE ') then
      Habilitation := ' AND ' + MonHabilitation.LeSQL
    else
      Habilitation := MonHabilitation.LeSQL;
    Where := Where + Habilitation;
  end;
  //FIN PT6}
  //DEB PT10
  if Where <> '' then
    Where := Where + ' AND PSA_SALARIE IN (SELECT DISTINCT PEF_SALARIE FROM ENFANTSALARIE)'
  else
    Where := ' WHERE PSA_SALARIE IN (SELECT DISTINCT PEF_SALARIE FROM ENFANTSALARIE)';
  //FIN PT10

  TFQRS1(Ecran).WhereSQL := SQL + Where + Tri;
end;

procedure TOF_PGEDIT_ENFANTS.OnLoad;
begin
  inherited;
end;

procedure TOF_PGEDIT_ENFANTS.OnArgument(S: string);
var
  Num: Integer;
  Defaut, Edit: THEdit;
  TCDefaut,Check : TCheckBox;
  Min, Max: string;
begin
  inherited;
  TriCN := '';
  for Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    if Num > 4 then Break;
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT'));
  Defaut := THEdit(GetControl('PSA_TRAVAILN1'));
  SetControlVisible('PSA_TRAVAILN1_', Defaut.Visible);
  SetControlVisible('CN1', Defaut.Visible);
  Defaut := THEdit(GetControl('PSA_TRAVAILN2'));
  SetControlVisible('PSA_TRAVAILN2_', Defaut.Visible);
  SetControlVisible('CN2', Defaut.Visible);
  Defaut := THEdit(GetControl('PSA_TRAVAILN3'));
  SetControlVisible('PSA_TRAVAILN3_', Defaut.Visible);
  SetControlVisible('CN3', Defaut.Visible);
  Defaut := THEdit(GetControl('PSA_TRAVAILN4'));
  SetControlVisible('PSA_TRAVAILN4_', Defaut.Visible);
  SetControlVisible('CN4', Defaut.Visible);
  Defaut := THEdit(GetControl('PSA_CODESTAT'));
  SetControlVisible('PSA_CODESTAT_', Defaut.Visible);
  SetControlVisible('CN5', Defaut.Visible);
  TCDefaut := TCheckBox(GetControl('CN1'));
  if TCDefaut <> nil then TCDefaut.OnClick := ChangeLieuTrav;
  TCDefaut := TCheckBox(GetControl('CN2'));
  if TCDefaut <> nil then TCDefaut.OnClick := ChangeLieuTrav;
  TCDefaut := TCheckBox(GetControl('CN3'));
  if TCDefaut <> nil then TCDefaut.OnClick := ChangeLieuTrav;
  TCDefaut := TCheckBox(GetControl('CN4'));
  if TCDefaut <> nil then TCDefaut.OnClick := ChangeLieuTrav;
  TCDefaut := TCheckBox(GetControl('CN5'));
  if TCDefaut <> nil then TCDefaut.OnClick := ChangeLieuTrav;
  Defaut := THEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    //  Defaut.text:=V_PGI_env.LibDossier;  //PT1 Mise en commentaire
    Defaut.text := GetParamSoc('SO_LIBELLE');
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
  // Debut PT2
  TFQRS1(Ecran).Caption := 'Liste Personnes à charge par salarié';   //PT11
  UpdateCaption(TFQRS1(Ecran));
  // FIN PT2
  Check := TCheckBox(GetControl('CKSORTIE'));
If Check <> Nil Then Check.OnCLick := CkSortieClick;
end;

procedure TOF_PGEDIT_ENFANTS.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PGEDIT_ENFANTS.ChangeLieuTrav(Sender: TObject);
var
  CN1, CN2, CN3, CN4, CN5: TCheckBox;
  LN1, LN2, LN3, LN4, LN5: THLabel;
begin
  TriCN := '';
  SetControlText('XX_RUPTURE1', '');
  SetControlText('XX_VARIABLE1', '');
  CN1 := TCheckBox(GetControl('CN1'));
  CN2 := TCheckBox(GetControl('CN2'));
  CN3 := TCheckBox(GetControl('CN3'));
  CN4 := TCheckBox(GetControl('CN4'));
  CN5 := TCheckBox(GetControl('CN5'));
  LN1 := THLabel(GetControl('TPSA_TRAVAILN1'));
  LN2 := THLabel(GetControl('TPSA_TRAVAILN2'));
  LN3 := THLabel(GetControl('TPSA_TRAVAILN3'));
  LN4 := THLabel(GetControl('TPSA_TRAVAILN4'));
  LN5 := THLabel(GetControl('TPSA_CODESTAT'));
  if (CN1 <> nil) and (CN2 <> nil) and (CN3 <> nil) and (CN4 <> nil) and (CN5 <> nil) then
  begin
    CN1.Enabled := True;
    CN2.Enabled := True;
    CN3.Enabled := True;
    CN4.Enabled := True;
    CN5.Enabled := True;
    if CN1.Checked = True then
    begin
      SetControlText('XX_RUPTURE1', 'PSA_TRAVAILN1');
      SetControltext('XX_VARIABLE1', LN1.Caption);
      TriCN := ' ORDER BY PSA_TRAVAILN1,';
      CN2.Enabled := False;
      CN3.Enabled := False;
      CN4.Enabled := False;
      CN5.Enabled := False;
    end;
    if CN2.Checked = True then
    begin
      SetControlText('XX_RUPTURE1', 'PSA_TRAVAILN2');
      SetControltext('XX_VARIABLE1', LN2.Caption);
      TriCN := ' ORDER BY PSA_TRAVAILN2,';
      CN1.Enabled := False;
      CN3.Enabled := False;
      CN4.Enabled := False;
      CN5.Enabled := False;
    end;
    if CN3.Checked = True then
    begin
      SetControlText('XX_RUPTURE1', 'PSA_TRAVAILN3');
      SetControltext('XX_VARIABLE1', LN3.Caption);
      TriCN := ' ORDER BY PSA_TRAVAILN3,';
      CN1.Enabled := False;
      CN2.Enabled := False;
      CN4.Enabled := False;
      CN5.Enabled := False;
    end;
    if CN4.Checked = True then
    begin
      SetControlText('XX_RUPTURE1', 'PSA_TRAVAILN4');
      SetControltext('XX_VARIABLE1', LN4.Caption);
      TriCN := ' ORDER BY PSA_TRAVAILN4,';
      CN1.Enabled := False;
      CN2.Enabled := False;
      CN3.Enabled := False;
      CN5.Enabled := False;
    end;
    if CN5.Checked = True then
    begin
      SetControlText('XX_RUPTURE1', 'PSA_CODESTAT');
      SetControltext('XX_VARIABLE1', LN5.Caption);
      TriCN := ' ORDER BY PSA_CODESTAT,';
      CN1.Enabled := False;
      CN2.Enabled := False;
      CN3.Enabled := False;
      CN4.Enabled := False;
    end;
  end;
end;

procedure TOF_PGEDIT_ENFANTS.CkSortieClick(Sender : TObject);
begin
        If Sender = Nil then Exit;
        If GetCheckBoxState('CKSORTIE') = CbChecked then
        begin
                SetControlVisible('DATESORTIE',True);
                SetControlText('DATESORTIE',DateToStr(Date));
                SetControlVisible('TDATEARRET',True);
        end
        else
        begin
                SetControlVisible('DATESORTIE',False);
                SetControlVisible('TDATEARRET',False);
        end;
end;


initialization
  registerclasses([TOF_PGEDIT_ENFANTS]);
end.

