{ Unité : Source TOF de la FICHE : TRMULFINPLAC
--------------------------------------------------------------------------------------
    Version   |   Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 0.91          20/10/03   JP   Création de l'unité
 6.0X.xxx.xxx  04/08/04   JP   Gestion des PopupMenus
 6.0.0.014.001 15/09/04   JP   FQ 10135 : sélection des écritures à supprimer
 6.51.001.001  15/11/05   JP   FQ 10307 : On court-circuite le onChange si en mode consultation
 7.05.001.001  23/10/06   JP   Gestion des filtres multi sociétés
--------------------------------------------------------------------------------------}
unit TRMULFINPLAC_TOF ;

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
  TOF_TRMULFINPLAC = class (TOF)
    procedure OnArgument(S : string); override;
  private
    {$IFDEF EAGLCLIENT}
    FListe : THGrid;
    {$ELSE}
    FListe : THDBGrid;
    {$ENDIF}
    PopupMenu : TPopUpMenu; {04/08/04}
    procedure DeleteOnClick(Sender : TObject);
    procedure InsertOnClick(Sender : TObject);
    procedure DupliqOnClick(Sender : TObject);
    procedure DeviseChange (Sender : TObject);
    procedure BanqueChange (Sender : TObject);
    procedure ListeDblClick(Sender : TObject);
  end ;

procedure TRLanceFiche_MulFinPlac(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  TomConditionFinPlac, Constantes;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_MulFinPlac(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULFINPLAC.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  TToolbarButton97(GetControl('BINSERT'   )).OnClick  := InsertOnClick;
  TToolbarButton97(GetControl('BDELETE'   )).OnClick  := DeleteOnClick;
  TToolbarButton97(GetControl('BDUPLIQUER')).OnClick  := DupliqOnClick;
  THValComboBox   (GetControl('BQ_DEVISE' )).OnChange := DeviseChange;
  THValComboBox   (GetControl('TCF_BANQUE')).OnChange := BanqueChange;
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
  THEdit(GetControl('TCF_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TCF_GENERAL')).DataType, '', '');
  SetPlusBancaire(THValComboBox(GetControl('TCF_BANQUE')), 'PQ', CODECOURANTS + ';' + CODETITRES + ';');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULFINPLAC.DeleteOnClick(Sender : TObject);
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
    SQL := 'DELETE FROM CONDITIONFINPLAC WHERE ' +
           'TCF_CONDITIONFP = "' + VarToStr(GetField('TCF_CONDITIONFP')) + '" AND ' +
           'TCF_GENERAL = "' + VarToStr(GetField('TCF_GENERAL')) + '" ';
    ExecuteSQL(SQL);
  end;

  {Pour raffraichir la liste}
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULFINPLAC.InsertOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_ConditionsFP('TR','TRCONDITIONSFP', '', '', 'ACTION=CREATION');
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULFINPLAC.DupliqOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_ConditionsFP('TR','TRCONDITIONSFP', '', '', 'ACTION=CREATION;' +
                                                           VarToStr(GetField('TCF_GENERAL')) + ';' +
                                                           VarToStr(GetField('TCF_CODECIB')) + ';' +
                                                           CODEDUPLICAT + ';');
  {Pour le rafraîchissement}
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULFINPLAC.ListeDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {15/11/05 : FQ 10307 : Mauvais appel à la fiche : il faut partir sur l'index}
  TRLanceFiche_ConditionsFP('TR','TRCONDITIONSFP', '', VarToStr(GetField('TCF_CONDITIONFP')) + ';' +
                                                       VarToStr(GetField('TCF_GENERAL')) + ';' {+
                                                       VarToStr(GetField('TCF_CODECIB'))}, 'ACTION=MODIFICATION');
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULFINPLAC.DeviseChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  AssignDrapeau(TImage(GetControl('IDEV')), THValComboBox(GetControl('BQ_DEVISE')).Value);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULFINPLAC.BanqueChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  SetControlText('TCF_GENERAL', '');
  if GetControlText('TCF_BANQUE') = '' then
    {23/10/06Gestion des filtres multi sociétés sur banquecp et dossier}
    THEdit(GetControl('TCF_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TCF_GENERAL')).DataType, '', '')
  else
    {23/10/06Gestion des filtres multi sociétés sur banquecp et dossier}
    THEdit(GetControl('TCF_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TCF_GENERAL')).DataType, '', '') +
                                              ' AND BQ_BANQUE = "' + THValComboBox(GetControl('TCF_BANQUE')).Value + '"';
end;

initialization
  RegisterClasses([TOF_TRMULFINPLAC]);

end.
