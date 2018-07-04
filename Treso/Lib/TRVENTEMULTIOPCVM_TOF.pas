{-------------------------------------------------------------------------------------
    Version   |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 6.20.xxx.xxx  15/12/04  JP   Création de l'unité
 7.09.001.001  23/10/06  JP   Gestion des filtres Multi sociétés
 7.09.001.001  15/12/06  JP   FQ 10388 : Violation d'accès lors du lancement du traitement si AllSelected
 7.09.001.003  21/12/06  JP   Refonte de la gestion CAMP
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}
unit TRVENTEMULTIOPCVM_TOF;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, mul, FE_Main,
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
  TOF_TRVENTEMULTIOPCVM = class (TOFCONF)
  {$ELSE}
  TOF_TRVENTEMULTIOPCVM = class (TOF)
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
  public
    CampOk   : Boolean; { = not GetParamSocSecur('SO_TRFIFO', False)}
    tVente   : TOB; {Tob contenant les OPCVM sélectionnés à vendre}

    procedure VenteMultiple   (Sender : TObject);

    {Constitue la tob de vente avec les enregistrements sélectionnés pour la fiche TRVENTEOPCVM}
    procedure PrepareTobVente ;
    procedure AjouteLigneVente(Q : THQuery = nil);
  end ;

procedure TRLanceFiche_VenteMultiOpcvm(Dom, Fiche, Range, Lequel, Arguments : string);


implementation

uses
  TRVENTEOPCVM_TOF, AglInit, HTB97, Constantes, Commun, ParamSoc, Math;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_VenteMultiOpcvm(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEMULTIOPCVM.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  Ecran.HelpContext := 50000139;
  {Affectation des évènements aux menus}
  GererPopup;
  {Autres évènements}
  TermineAffichage;
  {Tob contenant les OPCVM sélectionnés à vendre}
  tVente := TOB.Create('***', nil, -1);
  {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
  THValComboBox(GetControl('TOP_GENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TOP_GENERAL')).DataType, '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEMULTIOPCVM.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(tVente) then FreeAndNil(tVente);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEMULTIOPCVM.ListDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  tVente.ClearDetail;
  {Constitution de la tob de vente et affectation à TheTob}
  ChargeDetailOPCVM(tVente, GetField('TOP_PORTEFEUILLE'), GetField('TOP_CODEOPCVM'), GetField('TOP_GENERAL'));
  TheTOB := tVente;
  {Appel de la fiche de vente "Multiple"}
  TRLanceFiche_VenteOPCVM('TR', 'TRVENTEOPCVM', '', '', vop_Autre);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEMULTIOPCVM.GererPopup;
{---------------------------------------------------------------------------------------}
begin
  PopupMenu := TPopUpMenu(GetControl('POPUPMENU'));
  PopupMenu.Items[0].OnClick := VenteMultiple;
  AddMenuPop(PopupMenu, '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEMULTIOPCVM.TermineAffichage;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF EAGLCLIENT}
  TFMul(Ecran).bSelectAll.Visible := False;
  {$ENDIF}
  TFMul(Ecran).FListe.OnDblClick := ListDblClick;
  TFMul(Ecran).BOuvrir.OnClick   := ListDblClick;
  TFMul(Ecran).Binsert.Visible   := False;
  TToolbarButton97(GetControl('BMULTIPLE')).OnClick := VenteMultiple;
  SetControlVisible('BDELETE',  False);
end;

{Appel de l'écran de vente multiple
{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEMULTIOPCVM.VenteMultiple(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Teste s'il y a au moins une ligne de sélectionnée}
  if not TesteSelection then Exit;
  {Constitution de la tob de vente et affectation à TheTob}
  PrepareTobVente;
  if tVente.Detail.Count = 0 then 
    HShowMessage('0;' + Ecran.Caption + ';Toutes les parts ont déjà été vendue;I;O;O;O;', '', '')
  else begin
    {Appel de la fiche de vente "Multiple"}
    if CampOk then
      TRLanceFiche_VenteOPCVM('TR', 'TRVENTEOPCVM', '', '', vop_CAMP)
    else
      TRLanceFiche_VenteOPCVM('TR', 'TRVENTEOPCVM', '', '', vop_FIFO);
    {On vide la Tob de vente}
    tVente.ClearDetail;
  end;
  {Raffraîchissement du Mul}
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

{Constitue la tob de vente avec les enregistrements sélectionnés pour la fiche TRVENTEOPCVM
{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEMULTIOPCVM.PrepareTobVente;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  tVente.ClearDetail;

  {$IFNDEF EAGLCLIENT}
  TFMul(Ecran).Q.First;
  if TFMul(Ecran).FListe.AllSelected then
    while not TFMul(Ecran).Q.EOF do begin
      {On ne travaille que les opérations ouvertes
       15/12/06 : FQ 10388 : Le test est inutile car la vue filtre déjà sur TOP_STATUT <> "X"
      if (TFMul(Ecran).Q.FindField('TOP_STATUT').AsString <> 'X') then}
      AjouteLigneVente(TFMul(Ecran).Q);
      TFMul(Ecran).Q.Next;
    end
  else
  {$ENDIF}

  {On boucle sur la sélection}
  for n := 0 to TFMul(Ecran).FListe.nbSelected - 1 do begin
    TFMul(Ecran).FListe.GotoLeBookmark(n);
    {On ne travaille que les opérations ouvertes
    15/12/06 : FQ 10388 : Le test est inutile car la vue filtre déjà sur TOP_STATUT <> "X"
    if (VarToStr(GetField('TOP_STATUT')) <> 'X') then}
    AjouteLigneVente;
  end;
  TheTOB := tVente;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEMULTIOPCVM.AjouteLigneVente(Q : THQuery = nil);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  {On n'ajoute une ligne que si le nombre de parts vendues est strictement inférieur à
   celles achetées : théoriquement cela ne devrait pas arriver !!!}
  if GetValChp('TOP_NBPARTVENDU', Q) < GetValChp('TOP_NBPARTACHETE', Q) then begin
    {En Camp, on affiche les transactions de la ligne sélectionnée}
    T := TOB.Create('£££', tVente, -1);
    T.AddChampSupValeur('TOP_PORTEFEUILLE', GetValChp('TOP_PORTEFEUILLE', Q));
    T.AddChampSupValeur('TOP_CODEOPCVM'   , GetValChp('TOP_CODEOPCVM'   , Q));
    T.AddChampSupValeur('TOP_TRANSACTION' , '');
    T.AddChampSupValeur('TOP_GENERAL'     , GetValChp('TOP_GENERAL'     , Q));
    T.AddChampSupValeur('TOP_LIBELLE'     , 'Vente multiple');
    T.AddChampSupValeur('TOP_BASECALCUL'  , 0);
    T.AddChampSupValeur('TOP_MONTANTACH'  , GetValChp('TOP_MONTANTACH'  , Q));
    T.AddChampSupValeur('TOP_FRAISACH'    , GetValChp('TOP_FRAISACH'    , Q));
    T.AddChampSupValeur('TOP_TVAACHAT'    , GetValChp('TOP_TVAACHAT'    , Q));
    T.AddChampSupValeur('TOP_COMACHAT'    , GetValChp('TOP_COMACHAT'    , Q));
    T.AddChampSupValeur('TOP_COTATIONACH' , 0.0);
    T.AddChampSupValeur('TOP_NBPARTACHETE', GetValChp('TOP_NBPARTACHETE', Q));
    T.AddChampSupValeur('TOP_NBPARTVENDU' , GetValChp('TOP_NBPARTVENDU' , Q));
    {Champs "calculés"}
    T.AddChampSupValeur('MONTANTTOT', T.GetDouble('TOP_MONTANTACH') + T.GetDouble('TOP_FRAISACH') + T.GetDouble('TOP_TVAACHAT') + T.GetDouble('TOP_COMACHAT'));
    T.AddChampSupValeur('DATEVENTE', TDateTime(Max(V_PGI.DateEntree, TDateTime(GetValChp('TOP_DATEACHAT', Q)))));
    T.AddChampSupValeur('PARTAVENDRE', 0);
    T.AddChampSupValeur('MONTANTVEN', 0.0);
    T.AddChampSupValeur('TOP_DATEACHAT' , GetValChp('TOP_DATEACHAT', Q));
  end;
end;

{Avec le AllSelected, le GetField est inéficient : d'où cette petite fonction
{---------------------------------------------------------------------------------------}
function TOF_TRVENTEMULTIOPCVM.GetValChp(Chp : string; Q : THQuery = nil) : Variant;
{---------------------------------------------------------------------------------------}
begin
  if Q = nil then Result := GetField(Chp)
             else Result := Q.FindField(Chp).AsVariant;
end;

{Vérifie si au moins une ligne est sélectionnée
{---------------------------------------------------------------------------------------}
function TOF_TRVENTEMULTIOPCVM.TesteSelection : Boolean;
{---------------------------------------------------------------------------------------}
begin
  CampOk := not GetParamSocSecur('SO_TRFIFO', True);
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

initialization
  RegisterClasses([TOF_TRVENTEMULTIOPCVM]);

end.

