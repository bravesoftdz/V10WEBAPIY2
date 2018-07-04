{ Unité : Source TOF de la FICHE : TRMULEQUILIBRAGE
--------------------------------------------------------------------------------------
    Version    |   Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 0.91           20/10/03   JP    Création de l'unité
 6.0X.xxx.xxx   04/08/04   JP    Gestion des PopupMenus
 6.0.0.014.001  15/09/04   JP    FQ 10135 : sélection des écritures à supprimer
 7.05.001.001   23/10/06   JP    Gestion des filtres multi sociétés
--------------------------------------------------------------------------------------}
unit TRMULEQUILIBRAGE_TOF ;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  StdCtrls, Controls, Classes, HTB97, Commun, ExtCtrls,
  {$IFNDEF EAGLCLIENT}
  mul, FE_Main, HDB,
  {$ELSE}
  eMul, MaineAGL,
  {$ENDIF}
  Forms, HCtrls, HMsgBox, UTOF, Menus;

type
  TOF_TRMULEQUILIBRAGE = class (TOF)
    procedure OnArgument(S : string); override;
  private
    {$IFDEF EAGLCLIENT}
    FListe : THGrid;
    {$ELSE}
    FListe : THDBGrid;
    {$ENDIF}
    PopupMenu : TPopUpMenu; {04/08/04}
    procedure AgenceChange (Sender : TObject);
    procedure DeleteOnClick(Sender : TObject);
    procedure InsertOnClick(Sender : TObject);
    procedure DupliqOnClick(Sender : TObject);
    procedure DeviseChange (Sender : TObject);
    procedure ListeDblClick(Sender : TObject);
  end ;

procedure TRLanceFiche_MulEquilibrage(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  TomConditionEqui, Constantes;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_MulEquilibrage(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULEQUILIBRAGE.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 500013;
  TToolbarButton97(GetControl('BINSERT'   )).OnClick  := InsertOnClick;
  TToolbarButton97(GetControl('BDELETE'   )).OnClick  := DeleteOnClick;
  TToolbarButton97(GetControl('BDUPLIQUER')).OnClick  := DupliqOnClick;
  THValComboBox   (GetControl('BQ_DEVISE' )).OnChange := DeviseChange;
  THValComboBox   (GetControl('TCE_AGENCE')).OnChange := AgenceChange;
  {$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FLISTE'));
  FListe.MultiSelect := True; {FQ 10135}
  {$ELSE}
  FListe := THDBGrid(GetControl('FLISTE'));
  FListe.MultiSelection := True; {FQ 10135}
  {$ENDIF}
  FListe.OnDblClick := ListeDblClick;

  PopupMenu := TPopUpMenu(GetControl('POPUPMENU')); {04/08/04}
  PopupMenu.Items[0].OnClick := InsertOnClick;
  PopupMenu.Items[1].OnClick := DupliqOnClick;
  PopupMenu.Items[2].OnClick := DeleteOnClick;
  AddMenuPop(PopupMenu, '', '');
  {23/10/06Gestion des filtres multi sociétés sur banquecp et dossier}
  THEdit(GetControl('TCE_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TCE_GENERAL')).DataType, '', '');
  SetPlusBancaire(THValComboBox(GetControl('TCE_AGENCE')), 'TRA', CODECOURANTS + ';' + CODETITRES + ';');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULEQUILIBRAGE.DeleteOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  n   : Integer;
begin
  if FListe.nbSelected = 0 then begin
    HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner les conditions à supprimer.;W;O;O;O;', '', '');
    Exit;
  end;

  if HShowMessage('0;' + Ecran.Caption + ';Êtes-vous sûr de vouloir supprimer les lignes sélectionnées ?;Q;YN;N;N;', '', '') = mrNo then Exit;
  {15/09/04 : FQ 10135 : On fonctionne sur les lignes sélectionnées et non plus la ligne en cours}
  for n := 0 to FListe.NbSelected - 1 do begin
    FListe.GotoLeBookmark(n);
    SQL := 'DELETE FROM CONDITIONEQUI WHERE ' +
           'TCE_GENERAL = "' + VarToStr(GetField('TCE_GENERAL')) + '" ';
    ExecuteSQL(SQL);
  end;
  {Pour raffraichir la liste}
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULEQUILIBRAGE.InsertOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_ConditionEqui('TR','TRCONDITIONEQUI', '', '', 'ACTION=CREATION');
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULEQUILIBRAGE.DupliqOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_ConditionEqui('TR','TRCONDITIONEQUI', '', '', 'ACTION=CREATION;' + VarToStr(GetField('TCE_GENERAL')) + ';' + CODEDUPLICAT + ';');
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULEQUILIBRAGE.ListeDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_ConditionEqui('TR','TRCONDITIONEQUI', '', VarToStr(GetField('TCE_GENERAL')), 'ACTION=MODIFICATION');
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULEQUILIBRAGE.DeviseChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  AssignDrapeau(TImage(GetControl('IDEV')), THValComboBox(GetControl('BQ_DEVISE')).Value);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULEQUILIBRAGE.AgenceChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if GetControlText('TCE_AGENCE') = '' then
    {23/10/06Gestion des filtres multi sociétés sur banquecp et dossier}
    THEdit(GetControl('TCE_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TCE_GENERAL')).DataType, '', '')
  else begin
    {23/10/06Gestion des filtres multi sociétés sur banquecp et dossier}
    THEdit(GetControl('TCE_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TCE_GENERAL')).DataType, '', '') +
                                              ' AND BQ_AGENCE = "' + THValComboBox(GetControl('TCE_AGENCE')).Value + '"';
  end;
  SetControlText('TCE_GENERAL', '');
end;

initialization
  RegisterClasses([TOF_TRMULEQUILIBRAGE]);

end.
