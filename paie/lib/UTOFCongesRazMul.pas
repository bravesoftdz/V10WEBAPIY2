{***********UNITE*************************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 12/05/2004
Modifié le ... :   /  /
Description .. : Gestion de la RAZ des mouvements de congés payés
Mots clefs ... : PAIE;CP
*****************************************************************}
{
PT1  : 29/06/2004 PH V_50 Lancement suppression du bouton Bdelete au lieu
                          BOuvrir
PT2  : 02/12/2004 JL V_60 FQ 11870 Correction si ETABLISSEMENT="TOUS"
PT3  : 07/10/2005 SB V_65 FQ 12637 Mise à jour DBliste suite modif AGL
PT4  : 17/10/2007 VG V_80 Mise à jour du planning unifié
}
unit UTOFCongesRazMul;

interface
uses Classes,
     SysUtils,
     Controls,
{$IFDEF EAGLCLIENT}
     eMul,
{$ELSE}
     Mul,
{$ENDIF}
     Utof,
     Hctrls,
     HQry,
     Hent1,
     HMsgBox,
     Htb97;


type
  TOF_CongesRazMul = class(TOF)
    procedure OnArgument(stArgument: string); override;
    procedure OnLoad; override;
  private
    Q_Mul: THQuery;
    LanceSuppr: Boolean;
    procedure ExitEdit(Sender: TObject);
    procedure ChangeTypeSuppr(Sender: TObject);
    procedure ActiveWhere;
    procedure ClickBOuvrir(Sender: TObject);
  end;

implementation
uses EntPaie,
     PGoutils2,
     PgPlanningUnifie,
     YRessource;
{ TOF_CongesRazMul }


procedure TOF_CongesRazMul.OnArgument(stArgument: string);
var
  Combo: THValComboBox;
  Edit: THEdit;
  Btn: TToolBarButton97;
begin
  inherited;
  PgiInfo('Avant toute suppression de mouvements congés payés,#13#10' +
    'nous vous conseillons d''effectuer une sauvegarde de la base.', Ecran.caption);
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  SetControlText('TYPESUPPR', 'RAZCP');
  Combo := THValComboBox(GetControl('TYPESUPPR'));
  if combo <> nil then Combo.OnChange := ChangeTypeSuppr;
  Edit := THEdit(GetControl('SALARIE'));
  if Edit <> nil then Edit.OnExit := ExitEdit;
  Edit := THEdit(GetControl('SALARIE_'));
  if Edit <> nil then Edit.OnExit := ExitEdit;
  Btn := TToolBarButton97(GetControl('BDelete')); //PT1
  if Btn <> nil then Btn.Onclick := ClickBOuvrir;
end;

procedure TOF_CongesRazMul.OnLoad;
begin
  inherited;
  ActiveWhere;
  if Q_Mul <> nil then
  begin
    if (GetControlText('TYPESUPPR') = 'RAZCP') and (TFMul(Ecran).DBListe  <> 'PGMULCPSAL') then
      // Q_Mul.Liste := 'PGMULCPSAL'
      TFMul(Ecran).SetDBListe('PGMULCPSAL' )
    else
      if (GetControlText('TYPESUPPR') = 'RAZGRP') and (TFMul(Ecran).DBListe  <> 'PGMULMVTCPPRISGRP') then
      TFMul(Ecran).SetDBListe('PGMULMVTCPPRISGRP');
     // Q_Mul.Liste := 'PGMULMVTCPPRISGRP';
  end;
end;

procedure TOF_CongesRazMul.ActiveWhere;
var
  St, StWhere, StEtab, Etab: string;
  MultiCombo : THMultiValComboBox;
begin
  if LanceSuppr then exit;
  StWhere := '';
  StEtab := '';
  if GetControlText('SALARIE') <> '' then
  begin
    if (GetControlText('TYPESUPPR') = 'RAZCP') then
      StWhere := 'AND PSA_SALARIE>="' + GetControlText('SALARIE') + '" '
    else
      if (GetControlText('TYPESUPPR') = 'RAZGRP') then
      StWhere := 'AND PCN_SALARIE>="' + GetControlText('SALARIE') + '" ';
  end;
  if GetControlText('SALARIE_') <> '' then
  begin
    if (GetControlText('TYPESUPPR') = 'RAZCP') then
      StWhere := StWhere + 'AND PSA_SALARIE<="' + GetControlText('SALARIE_') + '" '
    else
      if (GetControlText('TYPESUPPR') = 'RAZGRP') then
      StWhere := StWhere + 'AND PCN_SALARIE<="' + GetControlText('SALARIE_') + '" ';
  end;
  MultiCombo := THMultiValComboBox(GetControl('ETABLISSEMENT'));
  if (GetControlText('ETABLISSEMENT') <> '') and (MultiCombo.Tous <> True) then  //PT2
  begin
    StEtab := '';
    Etab := GetControlText('ETABLISSEMENT');
    while (Etab <> '') do
    begin
      St := ReadTokenSt(Etab);
      if (GetControlText('TYPESUPPR') = 'RAZCP') then
        StEtab := StEtab + ' PSA_ETABLISSEMENT="' + St + '" OR'
      else
        if (GetControlText('TYPESUPPR') = 'RAZGRP') then
        StEtab := StEtab + ' PCN_ETABLISSEMENT="' + St + '" OR';
    end;
    if (StEtab <> '') then StEtab := ' AND (' + Copy(StEtab, 1, Length(StEtab) - 3) + ')';
  end;

  if IsValidDate(GetControlText('DATEVALIDITE')) then
    if (StrToDate(GetControlText('DATEVALIDITE')) <> idate1900) and (GetControlText('TYPESUPPR') = 'RAZGRP') then
      StWhere := StWhere + 'AND PCN_DATEVALIDITE>="' + UsDateTime(StrToDate(GetControlText('DATEVALIDITE'))) + '" ';

  if IsValidDate(GetControlText('DATEVALIDITE_')) then
    if (StrToDate(GetControlText('DATEVALIDITE_')) <> Idate2099) and (GetControlText('TYPESUPPR') = 'RAZGRP') then
      StWhere := StWhere + 'AND PCN_DATEVALIDITE<="' + UsDateTime(StrToDate(GetControlText('DATEVALIDITE_'))) + '" ';

  SetControlText('XX_WHERE', StWhere + StEtab);
end;

procedure TOF_CongesRazMul.ChangeTypeSuppr(Sender: TObject);
begin
  if (GetControlText('TYPESUPPR') = 'RAZCP') and (TFMul(Ecran).DBListe <> 'PGMULCPSAL') then
  begin
    SetControlText('DATEVALIDITE', DateToStr(Idate1900));
    SetControlText('DATEVALIDITE_', DateToStr(Idate2099));
    SetControlEnabled('DATEVALIDITE', False);
    SetControlEnabled('DATEVALIDITE_', False);
    ActiveWhere;
    TFMul(Ecran).SetDBListe('PGMULCPSAL'); //PT3 TFMul(Ecran).DBListe := 'PGMULCPSAL'; 
  end
  else
    if (GetControlText('TYPESUPPR') = 'RAZGRP') and (TFMul(Ecran).DBListe <> 'PGMULMVTCPPRISGRP') then
  begin
    SetControlEnabled('DATEVALIDITE', True);
    SetControlEnabled('DATEVALIDITE_', True);
    ActiveWhere;
    TFMul(Ecran).SetDBListe('PGMULMVTCPPRISGRP'); //PT3 TFMul(Ecran).DBListe := 'PGMULMVTCPPRISGRP';
  end;
end;

procedure TOF_CongesRazMul.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_CongesRazMul.ClickBOuvrir(Sender: TObject);
var
  Salarie, st: string;
  i, ordre: integer;
begin
  if Q_Mul = nil then exit;
  { Gestion de la sélection de salarié }
  if (TFMul(Ecran).FListe.nbSelected > 0) or (TFMul(Ecran).FListe.AllSelected) then
    if PgiAsk('Vous avez selectionné des salariés. Voulez-vous supprimer les congés payés que pour ces salariés?', Ecran.caption) = mrYes then
    begin
      St := '';
      { Suppression CP pour salarié sélectionné }
      { Composition du clause WHERE pour limiter le mul à ces salariés }
      for i := 0 to TFMul(Ecran).FListe.NbSelected - 1 do
      begin
        {$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
        {$ENDIF}
        TFMul(Ecran).FListe.GotoLeBookmark(i);
        if (TFMul(Ecran).DBListe = 'PGMULCPSAL') then
        begin
          Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
          St := St + ' PSA_SALARIE="' + Salarie + '" OR';
        end
        else
        begin
          Salarie := TFmul(Ecran).Q.FindField('PCN_SALARIE').asstring;
          St := St + ' PCN_SALARIE="' + Salarie + '" OR';
        end;
      end;
      TFMul(Ecran).FListe.ClearSelected;
      if St <> '' then
      begin
        St := GetControlText('XX_WHERE') + ' AND (' + Copy(St, 1, Length(st) - 2) + ')';
        SetControlText('XX_WHERE', St);
      end;
      LanceSuppr := True;
      TFMul(Ecran).BCherche.Click;
      LanceSuppr := False;
    end;
  if (TFMul(Ecran).DBListe = 'PGMULCPSAL') then
  begin
    { Confirmation de la suppression des CP }
    if PgiAsk('Etes-vous sûr de vouloir supprimer la totalité des congés payés des salariés affichés?', Ecran.caption) = mrNo then exit;
    { Balayage de la Query pour suppression CP }
    Q_Mul.First;
    while not Q_Mul.EOF do
    begin
      Salarie := Q_Mul.FindField('PSA_SALARIE').AsString;
      ExecuteSql('DELETE FROM ABSENCESALARIE WHERE PCN_TYPEMVT="CPA" AND PCN_SALARIE="' + Salarie + '"');
      DeletePGYPL ('', GetYRS_GUID (Salarie, '', ''));               //PT4
      Q_Mul.Next;
    end;
    TFMul(Ecran).BCherche.Click;
  end
  else
    if (TFMul(Ecran).DBListe  = 'PGMULMVTCPPRISGRP') then
  begin
    { Confirmation de la suppression des CP pris }
    if PgiAsk('Etes-vous sûr de vouloir supprimer les congés payés pris des salariés affichés?', Ecran.caption) = mrNo then exit;
    { Balayage de la Query pour suppression CP pris }
    Q_Mul.First;
    while not Q_Mul.EOF do
    begin
      Salarie := Q_Mul.FindField('PCN_SALARIE').AsString;
      Ordre := Q_Mul.FindField('PCN_ORDRE').AsInteger;
      ExecuteSql('DELETE FROM ABSENCESALARIE WHERE PCN_TYPEMVT="CPA" AND PCN_SALARIE="' + Salarie + '" ' +
        'AND PCN_ORDRE=' + IntToStr(ordre) + ' AND PCN_TYPECONGE="PRI"');
      DeletePGYPL ('', GetYRS_GUID (Salarie, '', ''));                   //PT4
      Q_Mul.Next;
    end;
    TFMul(Ecran).BCherche.Click;
  end;
end;

initialization
  registerclasses([TOF_CongesRazMul]);
end.

