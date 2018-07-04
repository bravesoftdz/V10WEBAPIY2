{-------------------------------------------------------------------------------------
    Version   | Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 8.01.001.004  31/01/07  JP   Création de l'unité : Mul des relevés ETEBAC
                              Reprise d'une partie du code de dpTOFCETEBAC qui ne concernait pas l'import
                              des fichiers proprement dit mais l'intégration dans CRELBQE :
08.00.001.025  17/07/07  LG   FQ 20561 : gestion des etablissement ds les guides
08.00.001.025  18/07/07  JP   FQ 21037 : Raccourci sur le F10
08.00.001.025  18/07/07  JP   FQ 21114 : Gestion du AllSelected lors de la suppression
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
 8.10.001.004  14/08/07  JP   FQ 21031 : CJ préfère un autre état de visualisation que je branche ne PCL
08.10.006.001  28/11/07  JP   FQ TRESO 10537 : Demande SIC : Multi impressions des relevés sélectionnés 
--------------------------------------------------------------------------------------}
unit CPMULCETEBAC_TOF;

interface

uses
  Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  EdtREtat, mul, FE_Main, HDB, DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ELSE}
  UtileAgl, eMul, MaineAGL,
  {$ENDIF}
  {$IFDEF TRCONF}
  uLibConfidentialite,
  {$ELSE}
  UTOF,
  {$ENDIF TRCONF}
  SysUtils, uTob, HCtrls, HEnt1;

type
  TTypeFiche = (TFichInteg, TFichPointage, TFichFromDP, TFichMul);

  {$IFDEF TRCONF}
  TOF_CPMULCETEBAC = class (TOFCONF)
  {$ELSE}
  TOF_CPMULCETEBAC = class (TOF)
  {$ENDIF TRCONF}
  private
    FTOBErreur       : TOB;
    FStETABQ         : string;
    FStGUICHET       : string;
    FstNUMEROCOMPTE  : string;
    FStJournal       : string;
    FStEta           : string ;
    FTypeFiche       : TTypeFiche;  {le type de fenêtre appelante}
    FInErreur        : Integer;
    FInNumEtebac     : Integer;
    FListFichier     : string;      // Liste des fichiers selectionnés
    FFormKeyDown     : TKeyEvent;

    procedure InitControles;
    procedure InitEvenements;
    procedure SelectCrit(aTyp : Char);
    procedure SelectionPartiel(const Etab, Guichet, Cpte : string);
  protected
    procedure PrepareAffichage;
    procedure ValidInteg;
    procedure VisualiserRel;
    {Remarque : si pas de commentaires "produit" sur la nouvelle ergonomie de la gestion des
                erreurs, supprimer tout ce qui concerne BVISU de la fiche et de la TOF car
                devenu inutile.}
    procedure VisualiserEcr;
  public
    edEtabBq  : THedit;
    edGuichet : THEdit;
    edNumCpte : THEdit;

    {$IFDEF EAGLCLIENT}
    FListe : THGrid;
    {$ELSE}
    FListe : THDBGrid;
    {$ENDIF}
    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
    procedure OnUpdate              ; override;
    procedure OnDisplay             ; override;

    function CreateControl : Boolean;

    procedure GuichetClick   (Sender : TObject);
    procedure EtabBqClick    (Sender : TObject);
    procedure NumCpteClick   (Sender : TObject);
    procedure BInsertClick   (Sender : TObject);
    procedure BDeleteClick   (Sender : TObject);
    procedure ListeDblClick  (Sender : TObject);
    procedure SelectRIBClick (Sender : TObject);
    procedure SelectEtabClick(Sender : TObject);
    procedure BVisuClick     (Sender : TObject);
    procedure BValideClick   (Sender : TObject);
    procedure FormKeyDown    (Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure PPF11Popup     (Sender : TObject);
    procedure SlctAllClick   (Sender : TObject); {JP 18/07/07 : FQ 18/07/07}
  end;

procedure CPLanceFiche_MulCEtebac(Arguments : string);


implementation

uses
  {$IFDEF VER150} Variants,{$ENDIF}
  {$IFDEF COMPTA} ImRapInt,{$ENDIF}
  ULibPointage,
  UObjEtats, HMsgBox, AglInit, HTB97, Windows, dpTOFCETEBAC, Grids, LookUp, Menus, ed_Tools;

{---------------------------------------------------------------------------------------}
procedure CPLanceFiche_MulCEtebac(Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPMULCETEBAC', '', '', Arguments);
end;

{---------------------------------------------------------------------------------------}
function TOF_CPMULCETEBAC.CreateControl : Boolean;
{---------------------------------------------------------------------------------------}
begin
 FTOBErreur := TOB.Create('***', nil, -1);
 Result     := True;
end;

{- LG - 17/07/2007 - Fb 20561 - gestion de l'etablissment ds les guides etebac
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  lStArg  : string;
  lStType : string;
  lEdit : THValComboBox;
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  Ecran.HelpContext := 50000129;
  
  lStArg            := S;

  FStETABQ        := ReadTokenST(lStArg);
  FStGUICHET      := ReadTokenST(lStArg);
  FstNUMEROCOMPTE := ReadTokenST(lStArg);
  lStType         := ReadTokenST(lStArg);
  FStJournal      := ReadTokenST(lStArg);
  FStEta          := lStArg ;

  if lStType = 'N' then
   begin // FB 20561 - LG
    lEdit := THValComboBox(GetControl('ETABLISSEMENT')) ;
    if FStEta = 'T' then
     lEdit.ItemIndex := 0
      else
        if lEdit.Values.IndexOf(FStEta) > - 1 then
         lEdit.value := FStEta ;
   end
    else
     begin
      SetControlVisible('ETABLISSEMENT' , false) ;
      SetControlVisible('FE__HLabel1' , false ) ;
     end ;

  {Decode la fenêtre appelante}
       if lStType = 'D' then FTypeFiche := TFichFromDP
  else if lStType = 'N' then FTypeFiche := TFichInteg
  else if lStType = 'P' then FTypeFiche := TFichPointage
  else FTypeFiche := TFichMul;

  if (FTypeFiche = TFichFromDP) or (FTypeFiche = TFichInteg) then
    CreateControl;

  InitControles;
  InitEvenements;
  PrepareAffichage;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(FTOBErreur) then begin
    if FTOBErreur.Detail.Count > 0 then begin
      TObjEtats.GenereEtatTob(FTOBErreur, TraduireMemoire('Liste des fichiers non intégrés'));
    end;
    FreeAndNil(FTOBErreur);
  end;
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited;
{En liaison avec ValidInteg; et VisualiserEcr;
 Reste donc la mise ne place de l'intégration dans CRELBQE
 cf UTOFSAISIEETEBAC.RechercheRibAIntegrer  ligne 4765
 cf UTOFCRELBQE.RechercheRibAIntegrer  ligne 4316
 Remplacer les AGLLanceFiche('CP','RLVETEBAC','','',';;;I') ; par CPLanceFiche_ImportCEtebac dans la Compta
  et les autres AGLLanceFiche('CP','RLVETEBAC',...) par CPLanceFiche_MulCEtebac(
}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.OnDisplay;
{---------------------------------------------------------------------------------------}
begin
  inherited;
//  SetControlEnabled('BVISU', GetField('ERREUR') = 2);
end;

(*
{---------------------------------------------------------------------------------------}
procedure TOF_CETEBAC.HGRibGetCellCanvas(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState);
{---------------------------------------------------------------------------------------}
begin
   {BVISU}
 if ( csDestroying in Ecran.ComponentState ) then Exit;
 if (ARow < 1 ) then exit ;
 if TOB(FTOBEnteteEtebac.Detail[ARow-1]).GetValue('ERREUR') = 1 then
  Canvas.Font.Color := clFuchsia
   else
    if TOB(FTOBEnteteEtebac.Detail[ARow-1]).GetValue('ERREUR') = 2 then
     Canvas.Font.Color := clRed ;
    //  else
   //    if Canvas.Font.Color
   //    Canvas.Font.Color := clWindowText;


end;
*)

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.BInsertClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if IsComptaPCL then Exit;
  {Appel de la fiche d'intégration des relevés}
  CPLanceFiche_ImportCEtebac('');
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.BDeleteClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n  : Integer;
  Q : TQuery;
begin
  if (TFMul(Ecran).FListe.nbSelected = 0) and not TFMul(Ecran).FListe.AllSelected then begin
    PGIError(TraduireMemoire('Veuillez sélectionner une ligne'), Ecran.Caption);
    Exit;
  end;

  if PgiAsk(TraduireMemoire('Attention, si vous avez supprimer le fichier des relevés sélectionnés') + #13 +
            TraduireMemoire('et que vous puger les données de la table des relevés, les informations') + #13 +
            TraduireMemoire('des relevés sélectionnés seront définitivement perdues.') + #13#13 +
            TraduireMemoire('Souhaitez-vous abandonner le traitement ?')) = mrYes then Exit;
  if PGIAsk(TraduireMemoire('Êtes-vous sûr de vouloir supprimer le(s) relevé(s) sélectionné(s) ?'), Ecran.Caption) = mrYes then begin
    BeginTrans;
    try
      {18/07/07 : FQ 21114 : Gestion du AllSelected}
      if TFMul(Ecran).FListe.AllSelected then begin
        {$IFDEF EAGLCLIENT}
        Q := TFMul(Ecran).Q.TQ;
        for n := 0 to Q.Detail.Count - 1 do
          ExecuteSQL('DELETE FROM CETEBAC WHERE CET_NUMRELEVE = ' + Q.Detail[n].GetString('CET_NUMRELEVE'));
        {$ELSE}
        Q := TFMul(Ecran).Q;
        while not Q.EOF do begin
          {Suppression du relevé courant ...}
          ExecuteSQL('DELETE FROM CETEBAC WHERE CET_NUMRELEVE = ' + Q.FindField('CET_NUMRELEVE').AsString);
          Q.Next;
        end;
        {$ENDIF EAGLCLIENT}
      end
      else begin
        {On boucle sur la sélection}
        for n := 0 to TFMul(Ecran).FListe.nbSelected - 1 do begin
          TFMul(Ecran).FListe.GotoLeBookmark(n);
          {Suppression du relevé courant ...}
          ExecuteSQL('DELETE FROM CETEBAC WHERE CET_NUMRELEVE = ' + VarToStr(GetField('CET_NUMRELEVE')));
        end;
      end;
      CommitTrans;
    except
      on E : Exception do begin
        RollBack;
        PGIError(E.Message);
      end;
    end;
    {Rafraîchissement du mul}
    TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.ListeDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  VisualiserRel;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.BVisuClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if FTypeFiche = TFichInteg then begin
    VisualiserEcr;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.BValideClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if FTypeFiche = TFichInteg then
    ValidInteg
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.FormKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
var
  OldKey : Word;
begin
  OldKey := Key;
  Key    := 0;

  case OldKey of
    {JP 18/07/07 : FQ 21037 : Branchement du F10}
    VK_F10    : if FTypeFiche = TFichMul then ListeDblClick(GetControl('BOUVRIR'))
                                         else BValideClick(GetControl('BVALIDER'));
    VK_F11    : TPopupMenu(GetControl('PPF11')).Popup(Mouse.CursorPos.x, Mouse.CursorPos.y);
    Ord('A')  : if ssCtrl in Shift then TFMul(Ecran).bSelectAllClick(TFMul(Ecran).bSelectAll);
    Ord('N')  : if ssCtrl in Shift then BInsertClick(Sender);
    Ord('R')  : if ssCtrl in Shift then SelectCrit('R');
    Ord('B')  : if ssCtrl in Shift then SelectCrit('B');
    Ord('I')  : if ssCtrl in Shift then ListeDblClick(Sender);
    VK_DELETE : if ssCtrl in Shift then BDeleteClick(Sender);
  else
    Key := OldKey;
  end;

  FFormKeyDown(Sender, Key, Shift);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.PPF11Popup(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF COMPTA}
  TMenuItem(GetControl('MMECR')).Visible := False;//(FTypeFiche = TFichInteg);
  {$ELSE}
  TMenuItem(GetControl('MMECR')).Visible := False;
  {$ENDIF COMPTA}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.SelectEtabClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  SelectCrit('B');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.SelectRIBClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  SelectCrit('R');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.EtabBqClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {On récupère les différentes valeurs de CET_ETABBQ dans CETEBAC}
  LookupList((Sender as THEdit), TraduireMemoire('Liste des établissements bancaires'), 'CETEBAC', 'CET_ETABBQ',
             '', '', 'CET_ETABBQ', True, 0, 'SELECT DISTINCT CET_ETABBQ FROM CETEBAC');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.GuichetClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  wh : string;
begin
  {Filtre éventuel sur les établissements bancaires}
  if edEtabBq.Text <> '' then
    wh := 'CET_ETABBQ LIKE "' + edEtabBq.Text + '%"';
  if wh <> '' then wh := 'WHERE ' + wh;

  {On récupère les différentes valeurs de CET_GUICHET dans CETEBAC}
  LookupList((Sender as THEdit), TraduireMemoire('Liste des guichets'), 'CETEBAC', 'CET_GUICHET',
             'CET_ETABBQ', wh, 'CET_ETABBQ, CET_GUICHET', True, 0, 'SELECT DISTINCT CET_GUICHET, CET_ETABBQ FROM CETEBAC ' + wh);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.NumCpteClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  wh : string;
  co : string;
  lj : string;
begin
  {Filtre éventuel sur les établissements bancaires}
  if edEtabBq.Text <> '' then
    wh := 'CET_ETABBQ LIKE "' + edEtabBq.Text + '%"';
  {Filtre éventuel sur les guichets bancaires}
  if edGuichet.Text <> '' then begin
    if wh = '' then wh := 'CET_GUICHET LIKE "' + edGuichet.Text + '%"'
               else wh := wh + ' AND CET_GUICHET LIKE "' + edGuichet.Text + '%"';
  end;
  if wh <> '' then wh := 'WHERE ' + wh;

  co := '';
  lj := '';
  {$IFDEF TRCONF}
  co := TObjConfidentialite.GetWhereConf(V_PGI.User, tyc_Banque, GetExeContext);
  {$ENDIF TRCONF}
  if (co <> '') then begin
    lj := 'LEFT JOIN BANQUECP ON CET_ETABBQ = BQ_ETABBQ AND CET_GUICHET = BQ_GUICHET AND CET_NUMEROCOMPTE = BQ_NUMEROCOMPTE ';
    if wh = '' then wh := 'WHERE ' + co
               else wh := wh + ' AND (' + co + ') ';
    wh := lj + wh;
  end;

  {On récupère les différentes valeurs de CET_NUMEROCOMPTE dans CETEBAC}
  LookupList((Sender as THEdit), TraduireMemoire('Liste des comptes '), 'CETEBAC', 'CET_NUMEROCOMPTE',
             'CET_ETABBQ;CET_GUICHET;', wh, 'CET_ETABBQ, CET_GUICHET, CET_NUMEROCOMPTE', True, 0, 'SELECT DISTINCT CET_NUMEROCOMPTE, CET_ETABBQ, CET_GUICHET FROM CETEBAC ' + wh);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.InitControles;
{---------------------------------------------------------------------------------------}
begin
  edEtabBq  := (GetControl('CET_ETABBQ'      ) as THedit);
  edGuichet := (GetControl('CET_GUICHET'     ) as THEdit);
  edNumCpte := (GetControl('CET_NUMEROCOMPTE') as THEdit);

  SetControlVisible('BVALIDER', FTypeFiche = TFichInteg);
  {$IFDEF COMPTA}
  SetControlVisible('BVISU'   , False);//FTypeFiche = TFichInteg);
  {$ELSE}
  SetControlVisible('BVISU'   , False);
  {$ENDIF COMPTA}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.InitEvenements;
{---------------------------------------------------------------------------------------}
begin
  {Evènements "métiers"}
  (GetControl('BINSERT' ) as TToolbarButton97).OnClick := BInsertClick;
  (GetControl('BDELETE' ) as TToolbarButton97).OnClick := BDeleteClick;
  (GetControl('BOUVRIR' ) as TToolbarButton97).OnClick := ListeDblClick;
  (GetControl('BVALIDER') as TToolbarButton97).OnClick := BValideClick;
  edNumCpte.OnElipsisClick := NumCpteClick;
  edGuichet.OnElipsisClick := GuichetClick;
  edEtabBq .OnElipsisClick := EtabBqClick;
  {Evènements du mul surchargés}
  TFMul(Ecran).FListe.OnDblClick := ListeDblClick;
  FFormKeyDown := TFMul(Ecran).OnKeyDown;
  TFMul(Ecran).OnKeyDown := FormKeyDown;
  {Impossible de faire un "as" ici, car TPopup et TMenuItem ne descendent pas de TControl !!}
  TPopupMenu(GetControl('PPF11')).OnPopup := PPF11Popup;
  TMenuItem(GetControl('MMSUP')).OnClick := BDeleteClick;
  TMenuItem(GetControl('MMADD')).OnClick := BInsertClick;
  TMenuItem(GetControl('MMALL')).OnClick := TFMul(Ecran).bSelectAllClick;
  TMenuItem(GetControl('MMRIB')).OnClick := SelectRIBClick;
  TMenuItem(GetControl('MMETA')).OnClick := SelectEtabClick;
  TMenuItem(GetControl('MMECR')).OnClick := BVisuClick;
  TMenuItem(GetControl('MMREL')).OnClick := ListeDblClick;

  if IsComptaPCL then begin
    TMenuItem(GetControl('MMADD')).Visible := False;
    (GetControl('BINSERT' ) as TToolbarButton97).Visible := False;
  end;
  {JP 18/07/07 : FQ 21114 : Gestion du bouton BSelectAll}
  TToolbarButton97(GetControl('BSELECTALL')).OnClick := SlctAllClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.PrepareAffichage;
{---------------------------------------------------------------------------------------}
var
  wh : string;
begin
  {$IFDEF EAGLCLIENT}
  FListe := THGrid(TFMul(Ecran).FListe);
  {$ELSE}
  FListe := THDBGrid(TFMul(Ecran).FListe);
  {$ENDIF}
  if FTypeFiche = TFichInteg then begin
    wh := 'CET_TYPELIGNE = "01" ';
    if FStETABQ        <> '' then wh := wh + 'AND CET_ETABBQ = "' + FStETABQ + '" ';
    if FStGUICHET      <> '' then wh := wh + 'AND CET_GUICHET = "' + FStGUICHET + '" ';
    if FStNUMEROCOMPTE <> '' then wh := wh + 'AND CET_NUMEROCOMPTE = "' + FStNUMEROCOMPTE + '" ';
    SetControlText('XX_WHERE', wh);
  end;
end;

{JP 18/07/07 : FQ 21114 : Gestion du selectAll
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.SlctAllClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Fiche : TFMul;
begin
  Fiche := TFMul(Ecran);
  {$IFDEF EAGLCLIENT}
  if not Fiche.FListe.AllSelected then begin
    if not Fiche.FetchLesTous then Exit;
  end;
  {$ENDIF}
  Fiche.bSelectAllClick(nil);
end;


type
  TTemp = class
   StValue : string;
   StEta   : string ;
  end;

{- LG - 17/07/2007 - FB 20561 - gestion des etablissement ds les guides
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.ValidInteg;
{---------------------------------------------------------------------------------------}
var
 n : Integer;
 T : TTemp;

    {-------------------------------------------------------------------}
    function _ControleFichier(vNumEtebac : Integer; vStJournal : string) : Integer;
    {-------------------------------------------------------------------}
    var
      lQ : TQuery;
    begin
      Result := 0;
      if vStJournal = '' then Exit;
      lQ := OpenSQL('SELECT COUNT(CRL_COMPTEUR) N FROM CRELBQE WHERE ' +
                    'CRL_DATECOMPTABLE >= (SELECT MIN(CET_DATEOPERATION) N1 FROM CETEBAC WHERE CET_NUMRELEVE = "' + IntToStr(vNumEtebac) + '" ) AND ' +
                    'CRL_DATECOMPTABLE <= (SELECT MAX(CET_DATEOPERATION) N1 FROM CETEBAC WHERE CET_NUMRELEVE = "' + IntToStr(vNumEtebac) + '" ) AND ' +
                    'CRL_JOURNAL = "' + vStJournal + '" ' , True);
      if lQ.FindField('N').AsInteger <> 0 then Result := 1;
      Ferme(lQ) ;

      if Result <> 0 then Exit;
      lQ := OpenSQL('SELECT COUNT(E_NUMEROPIECE) N FROM ECRITURE WHERE ' +
                    'E_DATECOMPTABLE >= (SELECT MIN(CET_DATEOPERATION) N1 FROM CETEBAC WHERE CET_NUMRELEVE = "' + IntToStr(vNumEtebac) + '") AND ' +
                    'E_DATECOMPTABLE <= (SELECT MAX(CET_DATEOPERATION) N1 FROM CETEBAC WHERE CET_NUMRELEVE = "' + IntToStr(vNumEtebac) + '") AND ' +
                    'E_JOURNAL = "' + vStJournal + '" ', True);
      if lQ.FindField('N').AsInteger <> 0 then Result := 2;
      Ferme(lQ) ;
    end ;


    {-------------------------------------------------------------------}
    procedure _AjouteFichier;
    {-------------------------------------------------------------------}

        {-----------------------------------------------------}
        procedure _AjouteErreur(Err : Byte);
        {-----------------------------------------------------}
        begin
          with TOB.Create('', FTOBErreur, -1) do begin
            if Err = 1 then
              AddChampSupValeur('ERREUR', TraduireMemoire('Présence de mvts etebac sur la période'))
            else
              AddChampSupValeur('ERREUR', TraduireMemoire('Présence écritures sur la période'));
            AddChampSupValeur('CET_RIB', '"C"' + VarToStr(GetField('CET_ETABBQ')) + '-' +
                                                 VarToStr(GetField('CET_GUICHET')) + '-' +
                                                 VarToStr(GetField('CET_NUMEROCOMPTE')));
            AddChampSupValeur('CET_DATEOPERATION', GetField('CET_DATEOPERATION'));
            AddChampSupValeur('CET_NUMRELEVE', GetField('CET_NUMRELEVE'));
          end;
        end;

    begin
      FInErreur           := _ControleFichier( FInNumEtebac , FStJournal);

      if (FInErreur = 1) and (PGIASK(TraduireMemoire('Il existe déjà des mouvements etebac pour cette période,') + #10#13 +
                                     TraduireMemoire('voulez-vous quand même intégrer le fichier ?'), Ecran.Caption) = mrNo) then begin
        _AjouteErreur(1);
        Exit;
      end
      else if (FInErreur = 2) and (PGIASK(TraduireMemoire('Il existe déjà des écritures pour cette période,') + #10#13 +
                                          TraduireMemoire('voulez-vous quand même intégrer le fichier ?'), Ecran.Caption) = mrNo) then begin
        _AjouteErreur(2);
        Exit;
      end;

      if FListFichier = '' then
        FListFichier := IntToStr(FInNumEtebac)
      else
        FListFichier := FListFichier + ';' + IntToStr(FInNumEtebac);
    end;


begin
  FListFichier        := '';

  if (FListe.nbSelected = 0) and not Fliste.AllSelected then begin
    PGIError(TraduireMemoire('Veuillez sélectionner une ligne'), Ecran.Caption);
    Exit;
  end;

  {$IFNDEF EAGLCLIENT}
  TFMul(Ecran).Q.First;
  if FListe.AllSelected then
    while not TFMul(Ecran).Q.EOF do begin
       FInNumEtebac  := TFMul(Ecran).Q.FindField('CET_NUMRELEVE').AsInteger;
       _AjouteFichier;
      TFMul(Ecran).Q.Next;
    end
  else
  {$ENDIF}

  for n := 0 to FListe.nbSelected - 1 do begin
    FListe.GotoLeBookmark(n);
     FInNumEtebac  := ValeurI(GetField('CET_NUMRELEVE'));
     _AjouteFichier;
  end;

  T := TTemp.Create;
  T.StValue := FListFichier;
  if FTypeFiche = TFichInteg then
   T.StEta   := GetControltext('ETABLISSEMENT') ;
  TheData := T;

  {On ferme l'écran après l'intégration}
  Ecran.Close;
end;

{28/06/04 - LG - nouvelle fonction de visualisation des ecritures en conflit ds la table ecriture
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.VisualiserEcr;
{---------------------------------------------------------------------------------------}
{$IFDEF COMPTA}
var
 lQ           : TQuery ;
 lInNumEtebac : integer ;
 lLigneCR     : TCRPiece ; // defini ds ImRapInt
 lCR          : TList ;
{$ENDIF COMPTA}
begin
{$IFDEF COMPTA}
 lCR               := TList.Create ;
 if (FListe.nbSelected <> 1) then Exit ;
 FListe.GotoLeBookMark(0) ;
 lInNumEtebac     := GetField('CET_NUMRELEVE') ;

 lQ := OpenSql('SELECT E_NUMEROPIECE,MIN(E_DATECOMPTABLE) DT, SUM(E_DEBIT) D, SUM(E_CREDIT) C FROM ECRITURE ' +
               'WHERE E_DATECOMPTABLE >= ' +
               '( SELECT MIN(CET_DATEOPERATION) FROM CETEBAC WHERE CET_NUMRELEVE =' + IntToStr(lInNumEtebac) + ' ) AND ' +
               ' AND E_DATECOMPTABLE <=  ' +
               '( SELECT MAX(CET_DATEOPERATION) FROM CETEBAC WHERE CET_NUMRELEVE =' + IntToStr(lInNumEtebac) + ' ) ' +
               ' AND E_JOURNAL = "' + FStJournal + '" ' +
               'GROUP BY E_NUMEROPIECE', True);
 while not lQ.Eof do
  begin
   lLigneCR             := TCRPiece.Create ;
   lLigneCR.NumPiece    := lQ.FindField('E_NUMEROPIECE').AsInteger ;
   lLigneCR.Journal     := FStJournal ;
   lLigneCR.QualifPiece := 'N' ;
   lLigneCR.Date        := lQ.FindField('dt').asDateTime ;
   lLigneCR.Debit       := lQ.FindField('d').asFloat ;
   lLigneCR.Credit      := lQ.FindField('c').asFloat ;
   lCR.Add (lLigneCR);
   lQ.Next ;
  end ;

 if lCR.Count > 0 then
  AfficheCRIntegration (lCR);

 VideListe(lCR) ;
 ferme(lQ) ;
{$ENDIF COMPTA}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.VisualiserRel;
{---------------------------------------------------------------------------------------}
var
  N : string;
  S : string;
  P : Integer;
  T : TOB;
  F : TOB;
  Q : TQuery;
begin
  {Récupération du numéro de relevé à imprimer}
  N := VarToStr(GetField('CET_NUMRELEVE'));
  if N = '' then begin
    PGIError(TraduireMemoire('Impossible de récupérer le numéro du relevé'), Ecran.Caption);
    Exit;
  end;

  {Préparation de la tob d'impression}
  T := TOB.Create('****', nil, -1);
  try
    if FListe.AllSelected or (FListe.nbSelected > 0) then begin
      N := '';

      if FListe.AllSelected then begin
        TFMul(Ecran).Q.First;
        while not TFMul(Ecran).Q.EOF do begin
          N := N + TFMul(Ecran).Q.FindField('CET_NUMRELEVE').AsString + ',';
          TFMul(Ecran).Q.Next;
        end;
      end
      else
        for p := 0 to FListe.nbSelected - 1 do begin
          FListe.GotoLeBookmark(p);
          N := N + VarToStr(GetField('CET_NUMRELEVE')) + ',';
        end;
      System.Delete(N, Length(N), 1);
    end;

    {28/11/07 : FQ 10537 : Nouvel état pour gérer le multi relevés}
    s := 'SELECT CET_DATEOPERATION, CET_DATEVALEUR, CET_LIBELLE, CET_DEVISE, CET_LIBELLE1, ' +
         'CET_DEBIT+CET_DEBITDEV CET_DEB, CET_CREDIT+CET_CREDITDEV CET_CRED, CET_CODEAFB, ' +
         'CET_ETABBQ||"-"||CET_GUICHET||"-"||CET_NUMEROCOMPTE CET_RIB, CET_TYPELIGNE, CET_NUMRELEVE ' ;
    if not (ctxLanceur in V_PGI.PGIContexte) then
      s := s + ', BQ_LIBELLE FROM CETEBAC LEFT JOIN BANQUECP ON BQ_ETABBQ = CET_ETABBQ AND ' +
               'BQ_GUICHET = CET_GUICHET AND BQ_NUMEROCOMPTE = CET_NUMEROCOMPTE '
    else
      s := s + ', "" BQ_LIBELLE FROM CETEBAC ';
    s := s + 'WHERE CET_NUMRELEVE IN (' + N + ') ORDER BY CET_NUMRELEVE, CET_TYPELIGNE, CET_DATEOPERATION';

    Q := OpenSQL(s, True);
    try
      while not Q.EOF do begin
        F := TOB.Create('****', T, -1);
        if Q.FindField('CET_TYPELIGNE').AsString = '01' then 
          F.AddChampSupValeur('LIBELLE' , TraduireMemoire('Solde initial'))
        else if Q.FindField('CET_TYPELIGNE').AsString = '07' then
          F.AddChampSupValeur('LIBELLE' , TraduireMemoire('Solde final'))
        else
          F.AddChampSupValeur('LIBELLE' , Q.FindField('CET_LIBELLE'      ).AsString);

        F.AddChampSupValeur('CIB'       , Q.FindField('CET_CODEAFB'      ).AsString);
        F.AddChampSupValeur('LIBELLEENR', Q.FindField('CET_LIBELLE1'     ).AsString);

        F.AddChampSupValeur('DATEVALEUR', Q.FindField('CET_DATEVALEUR'   ).AsString);
        F.AddChampSupValeur('DATERELEV' , Q.FindField('CET_DATEOPERATION').AsString);
        F.AddChampSupValeur('COMPTE'    , Q.FindField('BQ_LIBELLE'       ).AsString);
        F.AddChampSupValeur('BRIB'      , Q.FindField('CET_RIB'          ).AsString);
        F.AddChampSupValeur('TYPELIGNE' , Q.FindField('CET_TYPELIGNE'    ).AsString);
        F.AddChampSupValeur('DEVISE'    , Q.FindField('CET_DEVISE'       ).AsString);
        F.AddChampSupValeur('DEBIT'     , Q.FindField('CET_DEB'          ).AsFloat);
        F.AddChampSupValeur('CREDIT'    , Q.FindField('CET_CRED'         ).AsFloat);
        F.AddChampSupValeur('BRUPT'     , Q.FindField('CET_NUMRELEVE'    ).AsString);
        Q.Next;
      end;
    finally
      Ferme(Q);
    end;

    {28/11/07 : Suite à la FQ 21031, PCL avait un état différent de PGE, je suis parti de cet état que j'ai transformé en paysage}
    LanceEtatTob('E', 'ECT', 'RER', T, True, False, False, nil, '', TraduireMemoire('Relevé bancaires'), False);
  finally
    if Assigned(T) then FreeAndNil(T);
  end;
end;

{Sélection de tous les enregistrements qui correspondent soit au RIB ('R'), soit à
 l'établissement bancaire
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.SelectCrit(aTyp : Char);
{---------------------------------------------------------------------------------------}
var
  E : string;
  G : string;
  C : string;
begin
  {On ne fait de sélection partiel que s'il n'y pas plus d'une ligne sélectionée}
  if (Fliste.nbSelected > 1) or (Fliste.AllSelected) then Exit;

  E := VarToStr(GetField('CET_ETABBQ'));
  G := '';
  C := '';

  if aTyp = 'R' then begin
    G := VarToStr(GetField('CET_GUICHET'));
    C := VarToStr(GetField('CET_NUMEROCOMPTE'));
  end;
  SelectionPartiel(E, G, C);
end;

{$IFDEF EAGLCLIENT}
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.SelectionPartiel(const Etab, Guichet, Cpte : string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  k : Integer;
begin
  k := Fliste.Row;
  n := 1;
  TFmul(Ecran).Q.TQ.First;
  while not TFmul(Ecran).Q.TQ.EOF do begin
    if Etab = TFmul(Ecran).Q.TQ.FindField('CET_ETABBQ').AsString then begin
      if Guichet = '' then
        Fliste.FlipSelection(n)
      else if (Guichet = TFmul(Ecran).Q.TQ.FindField('CET_GUICHET').AsString) and
              (Cpte    = TFmul(Ecran).Q.TQ.FindField('CET_NUMEROCOMPTE').AsString) then
        Fliste.FlipSelection(n);
    end;
    TFmul(Ecran).Q.TQ.Next;
    Inc(n);
  end;
  Fliste.Row := k;
end;
{$ELSE}
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCETEBAC.SelectionPartiel(const Etab, Guichet, Cpte : string);
{---------------------------------------------------------------------------------------}
var
  SaveRow : TBookmark;
begin
  SaveRow := TFmul(Ecran).Q.GetBookmark;
  try
    TFmul(Ecran).Q.First;
    while not TFmul(Ecran).Q.Eof do begin
      if Etab = TFmul(Ecran).Q.FindField('CET_ETABBQ').AsString then begin
        if Guichet = '' then
          Fliste.FlipSelection
        else if (Guichet = TFmul(Ecran).Q.FindField('CET_GUICHET').AsString) and
                (Cpte    = TFmul(Ecran).Q.FindField('CET_NUMEROCOMPTE').AsString) then
          Fliste.FlipSelection;
      end;
      TFmul(Ecran).Q.Next;
    end;
  finally
    Fliste.Refresh;
    TFmul(Ecran).Q.GotoBookmark(SaveRow);
    TFmul(Ecran).Q.FreeBookmark(SaveRow);
  end;
end;
{$ENDIF EAGLCLIENT}

initialization
  RegisterClasses([TOF_CPMULCETEBAC]);

end.

