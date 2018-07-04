{-------------------------------------------------------------------------------------
    Version   |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 6.20.xxx.xxx  24/11/04  JP   Création de l'unité
 6.50.001.001  20/05/05  JP   Gestion de l'affichage des devises
 7.05.001.001  23/10/06  JP   Gestion des filtres Multi sociétés
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}
unit TRTRANSACOPCVMVEN_TOF;

interface

uses {$IFDEF VER150} variants,{$ENDIF} StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, Mul, FE_Main,
  {$ELSE}
  eMul, MaineAGL,
  {$ENDIF}
  Forms, SysUtils, HCtrls, HEnt1, HMsgBox, uTob,
  {$IFDEF TRCONF}
  uLibConfidentialite,
  {$ELSE}
  UTOF,
  {$ENDIF TRCONF}
  Menus, HQry;

type
  {$IFDEF TRCONF}
  TOF_TRTRANSACOPCVMVEN = class (TOFCONF)
  {$ELSE}
  TOF_TRTRANSACOPCVMVEN = class (TOF)
  {$ENDIF TRCONF}
    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
  private
    {Traitements sur l'interface utilisateur}
    procedure GererPopup;
    procedure TermineAffichage;
    function  TesteSelection  : Boolean;
    function  GetValChp       (Chp : string; Q : THQuery = nil) : Variant;
  protected
    {Composants et évènements du mul}
    PopupMenu : TPopUpMenu;

    procedure ListDblClick    (Sender : TObject);
    {20/05/05 : Gestion des devises en fonction du code OPCVM}
    procedure CodeOpcvmClick (Sender : TObject);
  public
    {Opérations de vente}
    tVente   : TOB; {Tob contenant les OPCVM sélectionnés à vendre}

    procedure VenteSimple     (Sender : TObject);
    procedure VentePartielle  (Sender : TObject);

    {Constitue la tob de vente avec les enregistrements sélectionnés pour la fiche TRVENTEOPCVM}
    procedure PrepareTobVente (TypeVent : string);
    procedure AjouteLigneVente(TypeVent : string; Q : THQuery = nil);
  end ;

procedure TRLanceFiche_TRansacOpcvmVen(Dom, Fiche, Range, Lequel, Arguments : string);


implementation

uses
  TROPCVM_TOM, TRVENTEOPCVM_TOF, AglInit, HTB97, Constantes, Math, Commun, ParamSoc;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_TRansacOpcvmVen(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRTRANSACOPCVMVEN.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  Ecran.HelpContext := 50000138;
  {Affectation des évènements aux menus}
  GererPopup;
  {Autres évènements}
  TermineAffichage;
  {Tob contenant les OPCVM sélectionnés à vendre}
  tVente := TOB.Create('***', nil, -1);
  {20/05/05 : Pour la gestion des devises}
  THValComboBox(GetControl('TOP_CODEOPCVM')).OnChange := CodeOpcvmClick;
  CodeOpcvmClick(GetControl('TOP_CODEOPCVM'));

  {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
  THValComboBox(GetControl('TOP_GENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TOP_GENERAL')).DataType, '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRTRANSACOPCVMVEN.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(tVente) then FreeAndNil(tVente);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRTRANSACOPCVMVEN.ListDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_OPCVM('TR', 'TROPCVM', '', GetField('TOP_NUMOPCVM'), ActionToString(taConsult));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRTRANSACOPCVMVEN.GererPopup;
{---------------------------------------------------------------------------------------}
begin
  PopupMenu := TPopUpMenu(GetControl('POPUPMENU'));
  PopupMenu.Items[0].OnClick := VenteSimple;
  PopupMenu.Items[1].OnClick := VentePartielle;
  AddMenuPop(PopupMenu, '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRTRANSACOPCVMVEN.TermineAffichage;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF EAGLCLIENT}
  TFMul(Ecran).bSelectAll.Visible := False;
  {$ENDIF}
  TFMul(Ecran).FListe.OnDblClick := ListDblClick;
  TFMul(Ecran).BOuvrir.OnClick   := ListDblClick;
  TFMul(Ecran).Binsert.Visible   := False;
  SetControlVisible('BSIMPLE', not GetParamSocSecur('SO_TRFIFO', True));
  TToolbarButton97(GetControl('BSIMPLE'  )).OnClick := VenteSimple;
  TToolbarButton97(GetControl('BPARTIEL' )).OnClick := VentePartielle;
  SetControlVisible('BDELETE',  False);
end;

{Traitement sur les lignes sélectionnées
{---------------------------------------------------------------------------------------}
procedure TOF_TRTRANSACOPCVMVEN.VenteSimple(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Teste s'il y a au moins une ligne de sélectionnée}
  if not TesteSelection then Exit;
  {Constitution de la tob de vente et affectation à TheTob}
  PrepareTobVente(vop_Simple);

  if tVente.Detail.Count = 0 then
    HShowMessage('0;' + Ecran.Caption + ';Toutes les parts ont déjà été vendue;I;O;O;O;', '', '')
  else begin
    {Appel de la fiche de vente "Multiple"}
    TRLanceFiche_VenteOPCVM('TR', 'TRVENTEOPCVM', '', '', vop_Simple);
    {On vide la Tob de vente}
    tVente.ClearDetail;
  end;
  {Raffraîchissement du Mul}
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

{Traitement sur la ligne sélectionnée et seulement elle
{---------------------------------------------------------------------------------------}
procedure TOF_TRTRANSACOPCVMVEN.VentePartielle(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if TFMul(Ecran).FListe.AllSelected or (TFMul(Ecran).FListe.nbSelected > 1) then
    HShowMessage('0;' + Ecran.Caption + ';Veuillez ne sélectionner qu''une ligne.;W;O;O;O;', '', '')

  else if TFMul(Ecran).FListe.nbSelected = 0 then
    HShowMessage('1;' + Ecran.Caption + ';Veuillez sélectionner une ligne.;W;O;O;O;', '', '')

  else if VarToStr(GetField('X')) = 'X' then
    HShowMessage('1;' + Ecran.Caption + ';Veuillez sélectionner une autre ligne,'#13 +
                 ' car cette transaction a été clôturée.;W;O;O;O;', '', '')

  else begin
    TRLanceFiche_OPCVM('TR', 'TROPCVM', '', GetField('TOP_NUMOPCVM'), ActionToString(taModif) + ';VEN;');
    TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
  end;
end;

{Constitue la tob de vente avec les enregistrements sélectionnés pour la fiche TRVENTEOPCVM
{---------------------------------------------------------------------------------------}
procedure TOF_TRTRANSACOPCVMVEN.PrepareTobVente(TypeVent : string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  tVente.ClearDetail;

  {$IFNDEF EAGLCLIENT}
  TFMul(Ecran).Q.First;
  if TFMul(Ecran).FListe.AllSelected then
    while not TFMul(Ecran).Q.EOF do begin
      {On ne travaille que les opérations ouvertes}
      if (TFMul(Ecran).Q.FindField('TOP_STATUT').AsString <> 'X') then
        AjouteLigneVente(TypeVent, TFMul(Ecran).Q);
      TFMul(Ecran).Q.Next;
    end
  else
  {$ENDIF}

  {On boucle sur la sélection}
  for n := 0 to TFMul(Ecran).FListe.nbSelected - 1 do begin
    TFMul(Ecran).FListe.GotoLeBookmark(n);
    {On ne travaille que les opérations ouvertes}
    if (VarToStr(GetField('TOP_STATUT')) <> 'X') then
      AjouteLigneVente(TypeVent);
  end;

  TheTOB := tVente;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRTRANSACOPCVMVEN.AjouteLigneVente(TypeVent : string; Q : THQuery = nil);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  {On n'ajoute une ligne que si le nombre de parts vendues est strictement inférieur à
   celles achetées : théoriquement cela ne devrait pas arriver !!!}
  if GetValChp('TOP_NBPARTVENDU', Q) < GetValChp('TOP_NBPARTACHETE', Q) then begin
    T := TOB.Create('£££', tVente, -1);
    T.AddChampSupValeur('TOP_NUMOPCVM'    , GetValChp('TOP_NUMOPCVM'    , Q));
    T.AddChampSupValeur('TOP_CODEOPCVM'   , GetValChp('TOP_CODEOPCVM'   , Q));
    T.AddChampSupValeur('TOP_TRANSACTION' , GetValChp('TOP_TRANSACTION' , Q));
    T.AddChampSupValeur('TOP_GENERAL'     , GetValChp('TOP_GENERAL'     , Q));
    T.AddChampSupValeur('TOP_LIBELLE'     , GetValChp('TOP_LIBELLE'     , Q));
    T.AddChampSupValeur('TOP_BASECALCUL'  , GetValChp('TOP_BASECALCUL'  , Q));
    T.AddChampSupValeur('TOP_MONTANTACH'  , GetValChp('TOP_MONTANTACH'  , Q));
    T.AddChampSupValeur('TOP_FRAISACH'    , GetValChp('TOP_FRAISACH'    , Q));
    T.AddChampSupValeur('TOP_TVAACHAT'    , GetValChp('TOP_TVAACHAT'    , Q));
    T.AddChampSupValeur('TOP_COMACHAT'    , GetValChp('TOP_COMACHAT'    , Q));
    T.AddChampSupValeur('TOP_COTATIONACH' , GetValChp('TOP_COTATIONACH' , Q));
    T.AddChampSupValeur('TOP_NBPARTACHETE', GetValChp('TOP_NBPARTACHETE', Q));
    T.AddChampSupValeur('TOP_NBPARTVENDU' , GetValChp('TOP_NBPARTVENDU' , Q));
    {Champs calculés}
    T.AddChampSupValeur('MONTANTTOT', T.GetDouble('TOP_MONTANTACH') + T.GetDouble('TOP_FRAISACH') + T.GetDouble('TOP_TVAACHAT') + T.GetDouble('TOP_COMACHAT'));
    T.AddChampSupValeur('DATEVENTE', TDateTime(Max(V_PGI.DateEntree, TDateTime(GetValChp('TOP_DATEACHAT', Q)))));
    T.AddChampSupValeur('PARTAVENDRE', GetValChp('TOP_NBPARTACHETE', Q) - GetValChp('TOP_NBPARTVENDU', Q));
    T.AddChampSupValeur('MONTANTVEN', 0.0);
    T.AddChampSupValeur('TOP_DATEACHAT' , GetValChp('TOP_DATEACHAT', Q));
  end;
end;

{Avec le AllSelected, le GetField est inéficient : d'où cette petite fonction
{---------------------------------------------------------------------------------------}
function TOF_TRTRANSACOPCVMVEN.GetValChp(Chp : string; Q : THQuery = nil) : Variant;
{---------------------------------------------------------------------------------------}
begin
  if Q = nil then Result := GetField(Chp)
             else Result := Q.FindField(Chp).AsVariant;
end;

{Vérifie si au moins une ligne est sélectionnée
{---------------------------------------------------------------------------------------}
function TOF_TRTRANSACOPCVMVEN.TesteSelection : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;
  {Aucune sélection, on sort}
  if (TFMul(Ecran).FListe.NbSelected = 0)
  {$IFNDEF EAGLCLIENT}
  and not TFMul(Ecran).FListe.AllSelected
  {$ENDIF}
  then begin
    HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner une ligne.;W;O;O;O;', '', '');
    Result := False;
  end;
end;

{20/05/05 : Affichage de la devise et du drapeau en fonction du code OPCVM
{---------------------------------------------------------------------------------------}
procedure TOF_TRTRANSACOPCVMVEN.CodeOpcvmClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  MajAffichageDevise(GetControl('IDEV'), GetControl('DEV'), GetControlText('TOP_CODEOPCVM'), sd_Opcvm);
end;

initialization
  RegisterClasses([TOF_TRTRANSACOPCVMVEN]);

end.

