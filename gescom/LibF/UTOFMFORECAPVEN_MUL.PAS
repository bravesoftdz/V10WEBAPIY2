{***********UNITE*************************************************
Auteur  ...... : O. TARCY                   
Créé le ...... : 28/12/2000                                                     
Modifié le ... : 23/07/2001
Description .. : Source TOF de la FICHE : RECAPVEN_MUL
Suite ........ : Lancement du récapitulatif par vendeur
Mots clefs ... : TOF;UTOFMFORECAPVEN_MUL;FO
*****************************************************************}
unit UTOFMFORECAPVEN_MUL;

interface
uses
  Classes, SysUtils,
  {$IFDEF EAGLCLIENT}
  eMul,
  {$ELSE}                                     
  Mul,
  {$ENDIF}
  HDB, HCtrls, HMsgBox, UTOF, HEnt1, ComCtrls, Forms, M3Fp, EntGC, HTB97
  ;

type
  TOF_MFORECAPVEN_MUL = class(TOF)
    public
      procedure OnArgument(S: string); override;
      procedure OnLoad ; override ;
      procedure FOVisuTicket;
      procedure FODebListe;
    private
      // JTR - Nvelles stats
      Stats : boolean;
      CbStats : THMultiValComboBox;
      Prefixe, TypeStats : string;
      Grid: THGrid;

      function TestSelection : boolean;
      procedure FListeOnDblClick(sender : TObject);
      procedure ExecEdition(Caisse : string; ZMin, ZMax : integer);
  end;

Const TexteMessage: array[1..4] of string = (
          {1}  'ERREUR DANS LA SELECTION',
          {2}  'Il faut sélectionner au moins une clôture',
          {3}  'Il faut sélectionner au moins une statistique',
          {4}  'Les clôtures sélectionnées doivent être consécutives'
              );

implementation
uses
  FOUtil, FODefi;

procedure TOF_MFORECAPVEN_MUL.OnArgument(S: string);
var stArg, Extract, Cond, Valeur: string;
    PComplement : TTabSheet;
begin
  inherited;
  stArg := S;
  Stats := False;
  Prefixe := '';
  TypeStats := '';
  while stArg <> '' do
  begin
    Extract := ReadTokenSt(stArg);
    Cond := copy(Extract, 1, pos('=',Extract)-1);
    Valeur := copy(Extract, pos('=',Extract)+1, length(Extract));
    if Cond = 'STATS' then
      Stats := (Valeur='TRUE')
    else if Cond = 'PREFIXE' then
      Prefixe := Valeur
    else if Cond = 'TYPE' then
      TypeStats := Valeur;
  end;
  Grid := THGRID(GetControl('FLISTE'));
  Grid.SetFocus;
  Grid.OnDblClick := FListeOnDblClick;
  THMultiValComboBox(GetControl('CHXSTATS')).Visible := Stats;
  THLabel(GetControl('T_CHXSTATS')).Visible := Stats;
  PComplement := TTabSheet(Ecran.FindComponent('PComplement'));
  PComplement.TabVisible := False;
  if Stats then
  begin
    CbStats := THMultiValComboBox(GetControl('CHXSTATS'));
    CbStats.DataType := 'GCTICKETSTATSG';
    if TypeStats = 'TABBORD' then
    begin
      CbStats.Plus := 'AND MO_NATURE="GSP"';
      Ecran.Caption := TraduireMemoire('Tableau de bord');
      Ecran.HelpContext := 113000502; // JTR - eQualité 11771
    end else
    if TypeStats = 'FAMILLE' then
    begin
      CbStats.Plus := 'AND MO_NATURE="GSP"';
      Ecran.Caption := TraduireMemoire('C.A. par familles d''articles');
    end else
    if TypeStats = 'PAIEMENT' then
    begin
      CbStats.Plus := 'AND MO_NATURE="GSE"';
      Ecran.Caption := TraduireMemoire('Mode de paiement');
      Ecran.HelpContext := 113000501; // JTR - eQualité 11771      
    end else
    if TypeStats = 'VENDEURS' then
    begin
      CbStats.Plus := 'AND MO_NATURE="GSV"';
      Ecran.Caption := TraduireMemoire('Vendeurs');
      Ecran.HelpContext := 113000503; // JTR - eQualité 11771
    end else
    if TypeStats = 'SPECIF' then
    begin
      CbStats.Plus := 'AND MO_NATURE="GSZ"';
      Ecran.Caption := TraduireMemoire('Statistiques spécifiques');
    end;
    if CbStats.Items.Count = 1 then
      CbStats.Text := CbStats.Values[0];
    UpdateCaption(Ecran);
  end else
    Ecran.HelpContext := 301100350; // JTR - eQualité 11771
  //Charge la caisse si vide
  if (not VH_GC.TOBPCaisse.FieldExists('GPK_CAISSE')) or (VH_GC.TOBPCaisse.GetValue('GPK_CAISSE') = '') then
    FOChargeVHGCCaisse(GetSynRegKey(CAISSEPARDEFAUT, '', True));
  SetControlText('GJC_DATEOUV', VH_GC.TOBPCaisse.GetValue('GPK_DATEOUV'));
  SetControlText('GJC_DATEOUV_', DateToStr(Date));
end;

procedure TOF_MFORECAPVEN_MUL.OnLoad ;
begin
  Inherited ;
  if GetControlText('GJC_CAISSE') = '' then
    SetControlText('GJC_CAISSE',VH_GC.TOBPCaisse.GetValue('GPK_CAISSE'));
end ;


procedure TOF_MFORECAPVEN_MUL.FOVisuTicket;
var F: TFMul;
  erreur: Boolean;
  st, sCaisse: string;
  nb, i, MinNumZ, MaxNumZ: Integer;
begin
  inherited;
  if not TestSelection then exit;
  F := TFMul(Ecran);
  nb := Grid.nbSelected;
  st := '';
  sCaisse := GetControlText('GJC_CAISSE');
  MinNumZ := 0;
  MaxNumZ := 0;
  if Grid.AllSelected then
  begin
    F.Q.First;
    while not F.Q.Eof do
    begin
      if (MinNumZ = 0) or (F.Q.FindField('GJC_NUMZCAISSE').asInteger < MinNumZ) then
        MinNumZ := F.Q.FindField('GJC_NUMZCAISSE').asInteger;
      if (MaxNumZ = 0) or (F.Q.FindField('GJC_NUMZCAISSE').asInteger > MaxNumZ) then
        MaxNumZ := F.Q.FindField('GJC_NUMZCAISSE').asInteger;
      F.Q.Next;
    end;
    F.bSelectAll.Click;
    Erreur := False;
  end else
  begin //vérification que les clôtures soient consécutives
    Grid.GotoLeBookmark(0);
    {$IFDEF EAGLCLIENT}
    F.Q.TQ.Seek(F.FListe.Row - 1);
    {$ENDIF}
    MinNumZ := F.Q.FindField('GJC_NUMZCAISSE').asInteger;
    MaxNumZ := F.Q.FindField('GJC_NUMZCAISSE').asInteger;
    for i := 1 to nb - 1 do
    begin
      Grid.GotoLeBookmark(i);
      {$IFDEF EAGLCLIENT}
      F.Q.TQ.Seek(F.FListe.Row - 1);
      {$ENDIF}
      if F.Q.FindField('GJC_NUMZCAISSE').asInteger < MinNumZ then
        MinNumZ := F.Q.FindField('GJC_NUMZCAISSE').asInteger else
        if F.Q.FindField('GJC_NUMZCAISSE').asInteger > MaxNumZ then
        MaxNumZ := F.Q.FindField('GJC_NUMZCAISSE').asInteger;
    end;
    if (MinNumZ = MaxNumZ) or (MaxNumZ - MinNumZ = nb - 1) then //1 seule clôture séctionnée ou clôtures consécutives
    begin
      Erreur := False;
    end else
    begin
      PGIInfo(TexteMessage[4], TexteMessage[1]);
      exit;
    end;
  end;

  if not Erreur then
  begin
    ExecEdition(sCaisse, MinNumZ, MaxNumZ);
    Grid.AllSelected := not Grid.AllSelected;
    TToolbarButton97(GetControl('bSelectAll')).Down := False;
    F.Q.First;
  end;
end;

procedure TOF_MFORECAPVEN_MUL.FODebListe;
begin
  TFMul(Ecran).Q.First;     
end;

procedure TOF_MFORECAPVEN_MUL.FListeOnDblClick(sender : TObject);
var sCaisse: string;
    NumZ : Integer;
begin                                          
  Grid.GotoLeBookmark(Grid.Row);
{$IFDEF EAGLCLIENT}
  TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
  sCaisse := GetControlText('GJC_CAISSE');
  NumZ := TFMul(Ecran).Q.FindField('GJC_NUMZCAISSE').asInteger;
  ExecEdition(sCaisse, NumZ, NumZ);
  Grid.AllSelected := not Grid.AllSelected;
  TToolbarButton97(GetControl('bSelectAll')).Down := False;
  TFMul(Ecran).Q.First;
end;

function TOF_MFORECAPVEN_MUL.TestSelection : boolean;
begin
  if (TFMul(Ecran).FListe.nbSelected = 0)
     and (not Grid.AllSelected)
     and (TToolbarButton97(GetControl('bSelectAll')).Down = False) then //Aucune lignes sélectionnées
  begin
    PGIInfo(TexteMessage[2], TexteMessage[1]);
    Result := False;
  end else
  if (Stats) and (CbStats.Text = '') then  // Aucune statistiques sélectionnées
  begin
    PGIInfo(TexteMessage[3], TexteMessage[1]);
    Result := False;
  end else
    Result := True;
end;

procedure TOF_MFORECAPVEN_MUL.ExecEdition(Caisse : string; ZMin, ZMax : integer);
var EtatSel, CodeEtat, Where: string;
begin
  Where := FOMakeWhereRecapVendeurs(Prefixe, Caisse, ZMin, ZMax);
  if stats then
  begin
    EtatSel := CbStats.Text;
    while EtatSel <> '' do
    begin
      CodeEtat := ReadTokenSt(EtatSel);
      if Prefixe = 'GPE' then
      begin
        Where := Where + ' AND MP_TYPEMODEPAIE<>"'+ TYPEPAIERESTEDU +'"';
        FOLanceImprimeLP(efoStatsGPE, Where, False, nil, CodeEtat);
      end else
      begin
        if TypeStats = 'VENDEURS' then
          FOLanceImprimeLP(efoStatsGPV, Where, False, nil, CodeEtat)
        else if ((TypeStats = 'TABBORD') or (TypeStats = 'FAMILLE')) then
          FOLanceImprimeLP(efoStatsGP, Where, False, nil, CodeEtat)
        else if TypeStats = 'SPECIF' then
          FOLanceImprimeLP(efoStatsGPZ, Where, False, nil, CodeEtat);
      end;
    end;
  end else
  begin
    FOLanceImprimeLP(efoRecapVend, Where, False, nil);
  end;
  Grid.AllSelected := not Grid.AllSelected;
  TToolbarButton97(GetControl('bSelectAll')).Down := False;
end;

initialization
  registerclasses([TOF_MFORECAPVEN_MUL]);
end.
