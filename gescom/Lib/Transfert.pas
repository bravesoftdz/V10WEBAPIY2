unit Transfert;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Facture, Menus, hmsgbox, HSysMenu, StdCtrls, ComCtrls, HRichEdt,
  HRichOLE, HTB97, Buttons, Grids, Hctrls, HFLabel, Mask, ExtCtrls, HPanel,
  FactUtil,HEnt1,UIUtil,EntGC,UTOB,Ent1,FactComm,M3FP,UtilFO,
{$IFDEF EAGLCLIENT}
{$ELSE}
  DBTables,
{$ENDIF}
  SaisUtil,StockUtil, ImgList, FactTOB, FactAdresse, FactPiece,    
  FactLotSerie, UtilPGI ;


//function  CreerTransfert(NaturePiece : String ) : boolean ;
function  CreerTransfert(NaturePiece, TransfertTEM_TRV : String) : boolean ;
//procedure SaisieTransfert( CleDoc : R_CleDoc ; Action : TActionFiche);
procedure SaisieTransfert( CleDoc : R_CleDoc ; Action : TActionFiche ; TransfertTEM_TRV : string ) ;  
Function  TransformeTRVenTRE(CleDoc : R_CleDoc): boolean;

type
  TFTransfert = class(TFFacture)
    HErrFO: THMsgBox;
    PEnteteFO: THPanel;
    HGP_NUMEROPIECE_: THLabel;
    FTitreTransfert: THLabel;
    GP_NUMEROPIECE_: THPanel;
    HGP_DATEPIECE_: THLabel;
    GP_DATEPIECE_: THCritMaskEdit;
    HGP_DEPOT_: THLabel;
    HGP_DEPOTDEST_: THLabel;
    GP_DEPOTDEST: THValComboBox;
    GP_DEPOT_: THValComboBox;
    HGP_ETABLISSEMENT_: THLabel;
    GP_ETABLISSEMENT_: THValComboBox;
    PCumul: THPanel;
    HGP_TOTALQTESTOCK_: THLabel;
    LGP_TOTALQTESTOCK_: THLabel;
    HGP_TOTALHTDEV_: THLabel;
    LGP_TOTALHTDEV_: THLabel;
    Procedure MAJPanelCumul;
    procedure GSEnter(Sender: TObject);
    procedure GP_ETABLISSEMENT_Change(Sender: TObject);
    procedure GP_DEPOTDESTChange(Sender: TObject);
    procedure GP_DEPOT_Change(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure PourlaligneClick(Sender: TObject); override;
    procedure PourledocumentClick(Sender: TObject); override;
  protected
    { Déclarations privées }
    DepotChange : boolean ;
    TobLitige, TF                 : Tob;
    DepotLitige, DepotDestination : string; 
    Procedure ShowDetail ( ARow : integer ) ; override;
    procedure InitPieceCreation ; override;
    function  CreationTRV : boolean;
    function  IsPieceExisteEtVivante(ChangeDateModif:boolean; var Vivante:boolean):boolean;
    //function  ConstruireTREouTRV(NouvelleNature : String; BSupp : boolean=False): boolean;                          
    function  ConstruireTREouTRV(NouvelleNature : String; BSupp : boolean=False; RefPrecedente : string=''): boolean; 
    procedure DetruitLaPiece ; override;
    procedure MAJCaisse;
    procedure CacherMontant(Cacher:boolean);
    procedure ClickDel ( ARow : integer ; AvecC,FromUser : boolean; SupDim: boolean=False; TraiteDim: boolean=False ) ; override;
    //procedure AppliqueTransfoDuplic ; override;
    procedure AppliqueTransfoDuplic(GestionAffichage: boolean=True) ; override;
    procedure ValideTransitEtLitige;                                           
  public
    { Déclarations publiques }
    procedure ClickValide (EnregSeul : Boolean=False) ; override ;
  end;

var
  FTransfert : TFTransfert;

implementation

uses
  FactSpec,
  ParamSoc,
  factcalc,
  TiersUtil,
  wcommuns,
  Math
  ;

{$R *.DFM}

function FindStInTab(Tab : array of string; St : string) : boolean;
var i : integer;
begin
  Result := False;
  for i:=Low(Tab) to High(Tab) do
    if UpperCase(Tab[i]) = UpperCase(St) then begin Result := True; Break; end;
end;

//Function CreerTransfert( NaturePiece : String ) : boolean ;
function  CreerTransfert(NaturePiece, TransfertTEM_TRV : String) : boolean ;  
Var CleDoc : R_CleDoc ;
    T_Auxi : String;
    IsTiersValide : Boolean;
begin
Result:=False ;
T_Auxi:=GetParamsoc('SO_GCTIERSDEFAUT');
IsTiersValide:=Not (T_Auxi='');
if IsTiersValide then IsTiersValide:=ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+T_Auxi+'"');
if Not IsTiersValide then
  begin PGIError('Vous devez renseigner un tiers valide dans les paramètres sociétés.',FTransfert.Caption); exit; end;
FillChar(CleDoc,Sizeof(CleDoc),#0) ;
CleDoc.NaturePiece:=NaturePiece ; CleDoc.DatePiece:=V_PGI.DateEntree ;
CleDoc.Souche:='' ; CleDoc.NumeroPiece:=0 ; CleDoc.Indice:=0 ;
//SaisieTransfert(CleDoc,taCreat);
SaisieTransfert(CleDoc,taCreat,TransfertTEM_TRV); 
Result:=True;
end;

//procedure SaisieTransfert( CleDoc : R_CleDoc ; Action : TActionFiche );
procedure SaisieTransfert( CleDoc : R_CleDoc ; Action : TActionFiche ; TransfertTEM_TRV : string ) ;  
var X : TFTransfert ;
    PP  : THPanel ;
begin
SourisSablier;
PP:=FindInsidePanel ;
X:=TFTransfert.Create(Application) ;
X.CleDoc:=CleDoc ; X.Action:=Action ;
if TransfertTEM_TRV='TEM_TRE' then //Transfert inter-dépôt sans logistique
  X.HelpContext := 110000313
else
  X.HelpContext := 110000424;
X.NewNature:=X.CleDoc.NaturePiece ; X.TransfoPiece:=False ; X.DuplicPiece:=False ;
X.TransfertTEM_TRV := TransfertTEM_TRV; 
X.Litige := False;
if PP=Nil then
   begin
    try
      X.ShowModal ;
    finally
      X.Free ;
    end ;
   SourisNormale ;
   end
else
   begin
   InitInside(X,PP) ;
   X.Show ;
   end;
end;

Function TransformeTRVenTRE(CleDoc : R_CleDoc): boolean;
var X : TFTransfert ;
    PP  : THPanel ;
begin
Result:=True ;
SourisSablier;
X:=TFTransfert.Create(Application) ;
X.CleDoc:=CleDoc ; X.Action:=taModif ;
X.NewNature:='TRE' ; X.TransfoPiece:=True ; X.DuplicPiece:=False ;
PP:=FindInsidePanel ;
if ((PP=Nil) or (ctxFO in V_PGI.PGIContexte)) then
   BEGIN
    try
      X.ShowModal ;
    finally
      X.Free ;
    end ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;


procedure TFTransfert.GSEnter(Sender: TObject);
Var bc,Cancel : boolean ;
    ACol,ARow : integer ;
begin
BeforeGSEnter(Self);
if Action=taConsult then Exit ;
if VH_GC.GCMultiDepots then
  if (GP_ETABLISSEMENT_.Text='') then begin HErrFO.Execute(5,Caption,''); exit; end;
if (GP_DEPOT_.Text='') or (GP_DEPOTDEST.Text='') then
  begin
  if VH_GC.GCMultiDepots then HErrFO.Execute(4,Caption,'') else HErrFO.Execute(1,Caption,'');
  if GP_DEPOT_.Text='' then GP_DEPOT_.SetFocus else GP_DEPOTDEST.SetFocus;
  exit;
  end;
//Met à jour l'entête de pièce
TOBPiece.PutValue('GP_DEPOT',GP_DEPOT_.Value);
if VH_GC.GCMultiDepots then TOBPiece.PutValue('GP_ETABLISSEMENT',GP_ETABLISSEMENT_.Value)
else TOBPiece.PutValue('GP_ETABLISSEMENT',GP_DEPOT_.Value);
TOBPiece.PutValue('GP_DATEPIECE',StrToDate(GP_DATEPIECE_.Text)); GP_DATEPIECE.Text:=GP_DATEPIECE_.Text;
TOBPiece.PutValue('GP_DEPOTDEST',GP_DEPOTDEST.Value);
// Bloque l'entête pour empêcher la modification du dépôt émetteur car tous les
// traitements sur les articles se font en fonction des dispos du dépôt émetteur
GP_ETABLISSEMENT_.Enabled:=False;
GP_DEPOT_.Enabled:=False; GP_DEPOTDEST.Enabled:=False;
//PEnteteFO.Enabled := False;
BeforeGSEnter(Self) ;
bc:=False ; Cancel:=False ; ACol:=GS.Col ; ARow:=GS.Row ;
GSRowEnter(GS,GS.Row,bc,False) ;
GSCellEnter(GS,ACol,ARow,Cancel) ;
EnabledGrid ; DejaRentre:=True ;
end;

// Met à jour le nombre d'articles transferés ainsi que le montant total de la pièce
Procedure TFTransfert.MAJPanelCumul;
begin
LGP_TOTALQTESTOCK_.Caption:=TobPiece.GetValue('GP_TOTALQTESTOCK');
if TobPiece.GetValue('GP_TOTALHTDEV')=0 then LGP_TOTALHTDEV_.Caption:='0,00'
else LGP_TOTALHTDEV_.Caption:=StrF00(TobPiece.GetValue('GP_TOTALHTDEV'),V_PGI.OkDecQ);
end;

Procedure TFTransfert.ShowDetail ( ARow : integer ) ;
begin
inherited;
MAJPanelCumul;
end;


procedure TFTransfert.InitPieceCreation ;
var Depot : string;
begin
//inherited;
CleDoc.Souche:=GetSoucheG(CleDoc.NaturePiece,VH^.EtablisDefaut,TOBPiece.GetValue('GP_DOMAINE')) ; CleDoc.NumeroPiece:=GetNumSoucheG(CleDoc.Souche) ;
GP_NUMEROPIECE.Caption:=HTitres.Mess[10] ;
InitTOBPiece(TOBPiece) ;
TOBPiece.PutValue('GP_NUMERO',CleDoc.NumeroPiece) ; TOBPiece.PutValue('GP_SOUCHE',CleDoc.Souche) ;
GP_DATEPIECE_.Text := GP_DATEPIECE.Text;
GP_NUMEROPIECE_.Caption:=GP_NUMEROPIECE.Caption;
if ctxFO in V_PGI.PGIContexte then GP_DATEPIECE_.Enabled:=False;
GeneCharge := False;
GP_TIERS.Text:= GetParamsoc('SO_GCTIERSDEFAUT') ;
ChargeTiers;
TiersVersAdresses(TOBTiers,TOBAdresses,TOBPiece);
GP_ETABLISSEMENT_.Enabled:=True;
GP_DEPOT_.Enabled:=True; GP_DEPOTDEST.Enabled:=True;
Depot := GP_DEPOT_.Value;
if (CtxMode in V_PGI.PGIContexte) and (Action=taCreat) then
  GP_DEPOT_.Plus := 'GDE_DEPOT<>"'+ GP_DEPOTDEST.Value + '" AND GDE_SURSITE="X"'
else
  GP_DEPOT_.Plus := 'GDE_DEPOT<>"'+ GP_DEPOTDEST.Value + '"';
GP_DEPOT_.Value := Depot;
DepotChange := True;
if VH_GC.GCMultiDepots then
  begin
  if (CtxMode in V_PGI.PGIContexte) and (Action=taCreat) then
    GP_ETABLISSEMENT_.Plus := 'ET_SURSITE="X"';
  GP_ETABLISSEMENT_.Value:=VH^.EtablisDefaut;
  GP_ETABLISSEMENT_.SetFocus;
  end
else
  begin
  GP_DEPOT_.Value := VH_GC.GCDepotDefaut;
  GP_DEPOTDEST.Plus := 'GDE_DEPOT<>"'+ GP_DEPOT.Value + '"';
  GP_DEPOT_.SetFocus;
  end;
MAJPanelCumul;
end;

function TFTransfert.IsPieceExisteEtVivante(ChangeDateModif:boolean; var Vivante:boolean):boolean;
var  Q : TQuery ;
begin
Result:=False; Vivante:=False;
Q:=OpenSQL('SELECT GP_DATEMODIF,GP_VIVANTE,GP_DATEPIECE FROM PIECE WHERE GP_NATUREPIECEG="'+CleDoc.NaturePiece+
'" AND GP_SOUCHE="'+CleDoc.Souche+
//'" AND GP_DATEPIECE="'+UsDateTime(CleDoc.DatePiece)+
'" AND GP_NUMERO='+IntToStr(CleDoc.NumeroPiece)+' AND GP_INDICEG='+IntToStr(CleDoc.Indice),True) ;
if not Q.Eof then
  begin
  if ChangeDateModif then
    begin
    TobPiece.PutValue('GP_DATEPIECE',Q.FindField('GP_DATEPIECE').AsDateTime);
    CleDoc.DatePiece := Q.FindField('GP_DATEPIECE').AsDateTime;
    TobPiece.PutValue('GP_DATEMODIF',Q.FindField('GP_DATEMODIF').AsDateTime);
    end;
  Vivante:=Q.FindField('GP_VIVANTE').Value='X';
  Result:=True;
  end;
Ferme(Q);
end;


// Permet de transformer la TOBPiece "Transfert Emis" en TOBPiece "Transfert Reçu"
// car le fonctionnement des transferts est basé sur 2 pièces.
//function TFTransfert.ConstruireTREouTRV(NouvelleNature : String; BSupp : boolean=False): boolean;
function  TFTransfert.ConstruireTREouTRV(NouvelleNature : String; BSupp : boolean=False; RefPrecedente : string=''): boolean;
var DepotDest,StMessage : string;
    PieceVivante : boolean;
    i : integer;
    Q : TQuery;
    lTransit : Boolean;
    DepotTransit : string;
    LignePrecedente : String;
begin
Result:=False; PieceVivante:=False;
if BSupp then StMessage:='supprimer' else StMessage:='modifier';
DepotTransit := '';                
if TobPiece.GetValue('GP_NATUREPIECEG')='TRE' then
begin
  DepotDest:=TobPiece.GetValue('GP_DEPOT'); TobPiece.PutValue('GP_NATUREPIECEG','TEM');
end
else
begin
  DepotDest:=TobPiece.GetValue('GP_DEPOTDEST'); TobPiece.PutValue('GP_NATUREPIECEG',NouvelleNature);

  if (TransfertTEM_TRV = 'TEM_TRV_NUMIDENTIQUE') or (TransfertTEM_TRV = 'TEM_TRV_NUMDIFFERENT') then
  begin
    lTransit := GetParamSoc('SO_GCTRV');
    if lTransit then
    begin
      DepotTransit := NullToVide(wGetSqlFieldValue('GDE_DEPOTTRANSIT', 'DEPOTS', 'GDE_DEPOT="' + TOBPiece.GetValue('GP_DEPOTDEST') + '"'));
      if DepotTransit <> '' then
        TobPiece.PutValue('GP_DEPOT', DepotTransit);
    end;
  end
  else
  begin
      TobPiece.PutValue('GP_DEPOT',DepotDest);
  end;
end;
CleDoc.NaturePiece:=TobPiece.GetValue('GP_NATUREPIECEG');
if Action=taModif then
  begin
  if TOBPiece.GetValue('GP_DATECREATION')<>Date then
    begin if Not DepotGererSurSite(DepotDest) then exit; end;

  if CleDoc.NaturePiece='TRV' then
    begin
    if IsPieceExisteEtVivante(True,PieceVivante) then
      begin
      if Not PieceVivante then
        begin
        TobPiece.PutValue('GP_NATUREPIECEG','TRE');
        CleDoc.NaturePiece:='TRE';
        end;
      end
    else
      begin PGIError('Impossible de trouver le transfert à valider.',Caption); exit; end;
    end;

  if TOBPiece.GetValue('GP_DATECREATION')<>Date then
    begin
    if mrNo=PGIAsk('Voulez-vous '+StMessage+' le '+RechDom('GCNATUREETAT',CleDoc.NaturePiece,False)+' ?',Caption) then
      exit;
    end;

  if Not PieceVivante then
    begin
    if Not IsPieceExisteEtVivante(True,PieceVivante) then
       begin
       PGIError('Impossible de trouver le '+RechDom('GCNATUREETAT',CleDoc.NaturePiece,False),Caption);
       exit;
       end;
    end;
  end;

// MAJ des lignes
For i:=0 to TOBPiece.Detail.Count-1 do
begin
  TobPiece.Detail[i].PutValue('GL_NATUREPIECEG',CleDoc.NaturePiece);
  if DepotTransit <> '' then
    TobPiece.Detail[i].PutValue('GL_DEPOT',DepotTransit)
  else                                                     
    TobPiece.Detail[i].PutValue('GL_DEPOT',DepotDest);
  TobPiece.Detail[i].PutValue('GL_ETABLISSEMENT',DepotDest);
{V500_004 Début}
  if (TransfertTEM_TRV = 'TEM_TRE') or (TransfertTEM_TRV = 'TEM_TRV_NUMDIFFERENT') then
    MajFromCleDoc(TOBPiece.Detail[i], CleDoc);
{V500_004 Fin}
  LignePrecedente := Copy(RefPrecedente, 1, Length(RefPrecedente) - 1)
  + IntToStr(TobPiece.detail[i].GetValue('GL_NUMORDRE')) + ';';
  TobPiece.Detail[i].PutValue('GL_PIECEPRECEDENTE', LignePrecedente);  
{V500_004 Début}
  if (action = taCreat) and (CleDoc.NaturePiece='TRV') and (TransfertTEM_TRV = 'TEM_TRV_NUMDIFFERENT') then
    TobPiece.Detail[i].PutValue('GL_VIVANTE', 'X');
{V500_004 Fin}
end;

// MAJ de PIEDBASE
For i:=0 to TOBBases.Detail.Count-1 do
begin
  TOBBases.Detail[i].PutValue('GPB_NATUREPIECEG',CleDoc.NaturePiece);
  if (TransfertTEM_TRV = 'TEM_TRE') or (TransfertTEM_TRV = 'TEM_TRV_NUMDIFFERENT') then
  begin
    TOBBases.Detail[i].PutValue('GPB_SOUCHE',CleDoc.Souche);
    TOBBases.Detail[i].PutValue('GPB_NUMERO',CleDoc.NumeroPiece);
    TOBBases.Detail[i].PutValue('GPB_INDICEG',CleDoc.Indice);
  end;
end;
// MAJ de PIEDECHE
For i:=0 to TOBEches.Detail.Count-1 do
begin
  TOBEches.Detail[i].PutValue('GPE_NATUREPIECEG',CleDoc.NaturePiece);
  if (TransfertTEM_TRV = 'TEM_TRE') or (TransfertTEM_TRV = 'TEM_TRV_NUMDIFFERENT') then
  begin
    TOBEches.Detail[i].PutValue('GPE_SOUCHE',CleDoc.Souche);
    TOBEches.Detail[i].PutValue('GPE_NUMERO',CleDoc.NumeroPiece);
    TOBEches.Detail[i].PutValue('GPE_INDICEG',CleDoc.Indice);
  end;
end;
// MAJ de PIEDPORT
For i:=0 to TOBPorcs.Detail.Count-1 do
begin
  TOBPorcs.Detail[i].PutValue('GPT_NATUREPIECEG',CleDoc.NaturePiece);
  if (TransfertTEM_TRV = 'TEM_TRE') or (TransfertTEM_TRV = 'TEM_TRV_NUMDIFFERENT') then
  begin
    TOBPorcs.Detail[i].PutValue('GPT_SOUCHE',CleDoc.Souche);
    TOBPorcs.Detail[i].PutValue('GPT_NUMERO',CleDoc.NumeroPiece);
    TOBPorcs.Detail[i].PutValue('GPT_INDICEG',CleDoc.Indice);
  end;
end;
// MAJ de LIGNELOT
For i:=0 to TOBDesLots.Detail.Count-1 do
begin
  TOBDesLots.Detail[i].PutValue('GLL_NATUREPIECEG',CleDoc.NaturePiece);
  if (TransfertTEM_TRV = 'TEM_TRE') or (TransfertTEM_TRV = 'TEM_TRV_NUMDIFFERENT') then
  begin
    TOBDesLots.Detail[i].PutValue('GLL_SOUCHE',CleDoc.Souche);
    TOBDesLots.Detail[i].PutValue('GLL_NUMERO',CleDoc.NumeroPiece);
    TOBDesLots.Detail[i].PutValue('GLL_INDICEG',CleDoc.Indice);
  end;
end;

if Action=taModif then
  begin
  // Lecture Pied
  Q:=OpenSQL('SELECT * FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False),True) ;
  TOBPiece_O.SelectDB('',Q) ;
  Ferme(Q) ;
  // Lecture Lignes
  Q:=OpenSQL('SELECT * FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,False)+' ORDER BY GL_NUMLIGNE',True) ;
  TOBPiece_O.LoadDetailDB('LIGNE','','',Q,False,True) ;
  Ferme(Q) ;
  PieceAjouteSousDetail(TOBPiece_O);
  // Lecture bases
  Q:=OpenSQL('SELECT * FROM PIEDBASE WHERE '+WherePiece(CleDoc,ttdPiedBase,False),True) ;
  TOBBases_O.LoadDetailDB('PIEDBASE','','',Q,False) ;
  Ferme(Q) ;
  // Lecture Echéances
  Q:=OpenSQL('SELECT * FROM PIEDECHE WHERE '+WherePiece(CleDoc,ttdEche,False),True) ;
  TOBEches_O.LoadDetailDB('PIEDECHE','','',Q,False) ;
  Ferme(Q) ;
  // Lecture Ports
  Q:=OpenSQL('SELECT * FROM PIEDPORT WHERE '+WherePiece(CleDoc,ttdPorc,False),True) ;
  TOBPorcs_O.LoadDetailDB('PIEDPORT','','',Q,False) ;
  Ferme(Q) ;
  // Lecture Lot
  LoadLesLots(TOBPiece_O,TOBLOT_O,CleDoc) ;
  end
else TOBPiece_O.Dupliquer(TOBPiece,True,True) ;

TOBPiece.PutValue('GP_DEVENIRPIECE','');
Result:=True;
end;

// C'est la même validation que dans la facture sauf que pour les transferts,
// on est obligé de valider 2 pièces différentes pour un même transfert
procedure TFTransfert.ClickValide (EnregSeul : Boolean=False) ;
Var io : TIOErr ;
    ResGC,i : integer ;
    RefPieceTEM,RefPieceTRE,NatureSuivante,NatureTemp,StLog : string;
    Ok,bForceEche,bOuvreEche : Boolean ;
    lTransit, lLitige               : Boolean;
    DepotTransit, Tiers, Depot      : string;
    CleDocTEM                       : R_CleDoc; 
begin
//Spécifique Transfert
if (NewNature='TRV') then exit;
if CtxFO in V_PGI.PGIContexte then MAJCaisse;
//Commun avec Facture
// Tests et actions préalables
if Not BValider.Enabled then Exit ;
if Action=taConsult then Exit ;
{$IFDEF MODE}
TraiteLaFusion(True);
{$ENDIF}
if Not SortDeLaLigne then Exit ;
if TOBGSAPleine then exit;
NextPrevControl(Self);
if (DuplicPiece) or ((GP_DEVISE.Value<>V_PGI.DevisePivot) and (Not EstMonnaieIN(GP_DEVISE.Value))) then
   begin
   PutValueDetail (TOBPiece,'GP_RECALCULER','X');
   CalculeLaSaisie(-1,-1,False) ;
   end;
//if (Action=taCreat) then TOBPiece.GetEcran(Self,PEntete);
if ctxAffaire in V_PGI.PGIContexte then TOBPiece.GetEcran(Self,PEnteteAffaire);
AfficheTaxes ; AffichePorcs ; PositionneVisa ;
if Not PieceModifiee then Exit ;
// Stat
if Not SaisieTablesLibres(TOBPiece) then Exit;
DepileTOBLignes(GS,TOBPiece,GS.Row,1) ;
if TesteRisqueTiers(True) then Exit ;
if TesteMargeMini(GS.Row) then Exit ;
// Gestion net à payer <> 0 si reglement obligatoire
if ((Action=taCreat) and (ObligeRegle)) then
   begin
     Case GCValideReglementOblig(TobPiece,TobTiers,TOBPieceRg,DEV) of
       0 : Exit;           {retour en saisie}
       1 : begin ClickAcompte(False) ; exit; end; {Retour en affectation de reglement}
       2 : ;               {on enregistre tout }
     end;
   end;
// Contrôle intégrité
ResGC:=GCPieceCorrecte(TOBPiece,TOBArticles,TOBCatalogu) ;
if ResGC>0 then BEGIN HErr.Execute(ResGC,Caption,'') ; Exit ; END ;
if ((Action=taModif) and (Not DuplicPiece)) then
  if Not PieceEncoreVivante(TOBPiece_O) then BEGIN HErr.Execute(7,Caption,'') ; Exit ; END ;
// Appels automatiques en fin de saisie
if Not BeforeValide(Self,TOBPiece,TOBBases,TOBTiers,TOBArticles,DEV,OldHT) then Exit ;
if OuvreAutoPort then ClickPorcs ;
if CompAnalP='AUT' then ClickVentil(True,False) ;

bForceEche:=((DEV.Code<>'') and (DEV.Code<>V_PGI.DevisePivot) and
             (DEV.Code<>V_PGI.DeviseFongible) and (Not EstMonnaieIn(DEV.Code))) ;
bOuvreEche:=(GereEche='AUT') or (bForceEche) ;
{$IFNDEF CHR}
if CtxMode in V_PGI.PGIContexte then Ok:=GereEcheancesMODE(TOBPiece,TOBTiers,TOBEches,TOBAcomptes,TOBPIECERG,Action,DEV,bOuvreEche)
                                else Ok:=GereEcheancesGC(TOBPiece,TOBTiers,TOBEches,TOBAcomptes,TOBPIECERG,Action,DEV,bOuvreEche) ;
if Not Ok then Exit else GP_MODEREGLE.Value:=TOBPiece.GetValue('GP_MODEREGLE') ;
{$ELSE}
if Not GereReglementsCHR(TOBPiece,TOBTiers,TOBEches,TOBAcomptes,Action,DEV,(GereEche='AUT')) then Exit else GP_MODEREGLE.Value:=TOBPiece.GetValue('GP_MODEREGLE') ;
{$ENDIF}

//Spécifique Transfert
// Enregistrement de la saisie
BValider.Enabled:=False ; ValideEnCours:=True ;
if ((NewNature='TRE') and TransfoPiece) then
begin
  DepotLitige := '';
  DepotDestination := TOBPiece.GetValue('GP_DEPOTDEST');       
  lTransit := GetParamSoc('SO_GCTRV') and (TobPiece_O.getvalue('GP_NATUREPIECEG') = 'TRV') and (NewNature='TRE');   
  if lTransit then                                             
    DepotTransit := NullToVide(wGetSqlFieldValue('GDE_DEPOTTRANSIT', 'DEPOTS', 'GDE_DEPOT="' + TOBPiece.GetValue('GP_DEPOTDEST') + '"')); 
  lTransit := lTransit and (DepotTransit <> '');

  if lTransit then                                             
    DepotLitige := NullToVide(wGetSqlFieldValue('GDE_DEPOTLITIGE', 'DEPOTS', 'GDE_DEPOT="' + TOBPiece.GetValue('GP_DEPOTDEST') + '"')); 

  lLitige := lTransit and (DepotLitige <> '');
  if lLitige then TobLitige := Tob.Create('_TEMP_', nil, -1);   

  For i:=0 to TOBPiece.Detail.Count-1 do
  begin
    TobPiece.Detail[i].PutValue('GL_DEPOT',TOBPiece.GetValue('GP_DEPOTDEST'));
    TobPiece.Detail[i].PutValue('GL_ETABLISSEMENT',TOBPiece.GetValue('GP_DEPOTDEST'));
    if lLitige and (TobPiece.detail[i].getvalue('GL_QTESTOCK') <> TobPiece_O.detail[i].getvalue('GL_QTESTOCK')) then 
    begin                                                       
      TF := Tob.Create('_TEMP_', TobLitige, -1);
      TF.AddChampSupValeur('INDICE',i);
    end;                                                        
  end;

  //Dans ce cas il faut conserver la pièce TRV et crééer une nouvelle pièce TRE
  if lTransit then                                              
  begin                                                         
    TransfertTEM_TRV := 'TEM_TRV_NUMDIFFERENT';
    TransfoPiece := True;
    V_PGI.IoError  := Transactions(ValideTransitEtLitige,1);
  end
  else
    V_PGI.IoError := Transactions(ValideLaPiece,1);

  if lLitige then TobLitige.free;                               

end
else
begin
  Transactions(ValideNumeroPiece,2) ;

  BEGINTRANS;

  RefPieceTEM:=EncodeRefPiece(TOBPiece,CleDoc.NumeroPiece); //Modif 06/06/2002

  //if CreationTRV then NatureSuivante:='TRV' else NatureSuivante:='TRE';
  if (TransfertTEM_TRV = 'TEM_TRE') then
  begin
    CleDocTEM := CleDoc;
    NatureSuivante:='TRE';
  end
  else
  begin                                     
    if (TransfertTEM_TRV = 'TEM_TRV_NUMDIFFERENT') then
      CleDocTEM := CleDoc;
    if CreationTRV then NatureSuivante:='TRV' else NatureSuivante:='TRE';
  end;                                                  

  if Action=taCreat then
  begin
    TobPiece.PutValue('GP_NATUREPIECEG',NatureSuivante);
    RefPieceTRE:=EncodeRefPiece(TOBPiece,CleDoc.NumeroPiece); //Modif 06/06/2002
    TobPiece.PutValue('GP_NATUREPIECEG','TEM');
    TOBPiece.PutValue('GP_DEVENIRPIECE',RefPieceTRE);
    Depot := TobPiece.GetValue('GP_DEPOT');
    Tiers := RechercheTiersDepot(Depot);
    if Tiers <> TobPiece.GetValue('GP_TIERS') then
    begin
      GP_Tiers.text := Tiers;
      ChargeTiers;
      TobPiece.P('GP_DEPOT', Depot);  //car ChargeTiers, IncidenceTiers (TOBPiece.GetEcran(Self)) change le depôt en fonction de l'écran ;
    end;
  end;
  if V_PGI.IoError=oeOk then ValideLaPiece;
  // Modifie la TOBPiece pour donner un Transfert à valider ou un transfert reçu
  // en fonction du paramètre société SO_GCTRV
  if V_PGI.IoError=oeOk then
    begin
    NatureTemp := CleDoc.NaturePiece;
    if (TransfertTEM_TRV = 'TEM_TRE') or (TransfertTEM_TRV = 'TEM_TRV_NUMDIFFERENT') then
    begin
      CleDoc.Souche := GetSoucheG(NatureSuivante, TOBPiece.GetValue('GP_ETABLISSEMENT'), TOBPiece.GetValue('GP_DOMAINE'));
      TobPiece.P('GP_SOUCHE', CleDoc.Souche);
      Transactions(ValideNumeroPiece,2) ;
      MajFromCleDoc(TOBPiece, CleDoc);
    end;
    //if ConstruireTREouTRV(NatureSuivante) then
    if ConstruireTREouTRV(NatureSuivante, False, RefPieceTEM) then  
      begin
      Depot := TobPiece.GetValue('GP_DEPOT');
      Tiers := RechercheTiersDepot(Depot);
      if Tiers <> TobPiece.GetValue('GP_TIERS') then
      begin
        GP_Tiers.text := Tiers;
        ChargeTiers;
        TobPiece.P('GP_DEPOT', Depot);              //car ChargeTiers, IncidenceTiers (TOBPiece.GetEcran(Self)) change le depôt en fonction de l'écran ;
      end;
{V500_004 Début}
      if (action = taCreat) and (CleDoc.NaturePiece='TRV') and (TransfertTEM_TRV = 'TEM_TRV_NUMDIFFERENT') then
        TobPiece.PutValue('GP_VIVANTE', 'X');
{V500_004 Fin}
      ValideLaPiece;
      if (TransfertTEM_TRV = 'TEM_TRE') or (TransfertTEM_TRV = 'TEM_TRV_NUMDIFFERENT') then
      begin
        RefPieceTRE := EncodeRefPiece(TOBPiece,CleDoc.NumeroPiece);
        ExecuteSQL('UPDATE PIECE SET GP_DEVENIRPIECE="' + RefPieceTRE + '" WHERE ' + WherePiece(CleDocTEM, ttdPiece , False));
      end;
      DecodeRefPiece(RefPieceTEM,CleDoc); TOBPiece.PutValue('GP_NATUREPIECEG',CleDoc.NaturePiece);
      end
    else
      begin
      CleDoc.NaturePiece := NatureTemp; TOBPiece.PutValue('GP_NATUREPIECEG',NatureTemp);
      end;
    end;
  if V_PGI.IoError=oeOk then COMMITTRANS
  else
    begin
    ROLLBACK;
    StLog := 'Erreur ROLLBACK validation des 2 transferts, Piece n°'+IntToStr(TOBPiece.GetValue('GP_NUMERO'));
    AjouteEvent(TOBPiece.GetValue('GP_NATUREPIECEG'),StLog,2);
    end;
end;
//MontreNumero(TOBPiece) ;
GP_NUMEROPIECE_.Caption:=GP_NUMEROPIECE.Caption;
BValider.Enabled:=True ;

Case V_PGI.IoError of
        oeOk : BEGIN
               ForcerFerme:=True ; AvoirDejaInverse:=False ;
               AfterValide(Self,TOBPiece,TOBBases,TOBTiers,TOBArticles,DEV);
               END ;
   oeUnknown : BEGIN MessageAlerte(HTitres.Mess[5])  ; ValideEnCours:=False ; Exit ; END ;
    oeSaisie : BEGIN MessageAlerte(HTitres.Mess[6])  ; ValideEnCours:=False ; Exit ; END ;
  oePointage : BEGIN MessageAlerte(HTitres.Mess[23]) ; ValideEnCours:=False ; Exit ; END ;
  oeLettrage : BEGIN MessageAlerte(HTitres.Mess[13]) ; ValideEnCours:=False ; Exit ; END ;
  oeStock    : BEGIN MessageAlerte(HTitres.Mess[27]) ; ValideEnCours:=False ; Exit ; END ;
     else  BEGIN MessageAlerte(HTitres.Mess[5])  ; ValideEnCours:=False ; Exit ; END ;
   END ;
BValider.Enabled:=False; ValideEnCours:=False ;
if (GetInfoParPiece(NewNature,'GPP_IMPIMMEDIATE')='X') or (GetInfoParPiece(NewNature,'GPP_VALMODELE')='X') or (GetInfoParPiece(NewNature,'GPP_IMPETIQ')='X')  then
   BEGIN
   if ((Action=taCreat) or (Action=taModif) or (DuplicPiece) or (TransfoPiece)) and
      Not(SaisieTypeAvanc) then
      BEGIN
      io:=Transactions(ValideImpression,1) ;
      if io<>oeOk then
        BEGIN
        StLog := 'Erreur ValideImpression, Piece n°'+IntToStr(TOBPiece.GetValue('GP_NUMERO'));
        AjouteEvent(TOBPiece.GetValue('GP_NATUREPIECEG'),StLog,3);
        DeflagEdit(TOBPiece) ;
        MessageAlerte(HTitres.Mess[19]) ;
        END ;
      END ;
   END ;
if Action<>taCreat then
   BEGIN
   if ((DuplicPiece) or (TransfoPiece)) then MontreNumero(TOBPiece) ;
   Close ;
   END else
   BEGIN
   VH_GC.GCLastRefPiece:=EncodeRefPiece(TOBPiece) ;
   MontreNumero(TOBPiece) ;
   if ((PasBouclerCreat) or (ctxBTP in V_PGI.PGIContexte)) then Close else ReInitPiece ;
   END ;
BValider.Enabled:=True;
end;

procedure TFTransfert.ValideTransitEtLitige;
var
  PiecePrecedente, Depot        : string;
  LignePrecedente               : string;
  Tiers, St                     : string;
  i, ResGC                      : integer;
  Q                             : TQuery;
  Ok                            : boolean;
  TOBP_Origine, TobF            : tob;
  TobQteReste, T1, T2           : tob;
  nQteStock, nQteFact           : Double;
  nQteTotalStock, nQteTotalFact : Double;

  procedure RazTobPourPRE;
  begin
    // Pièce
    TOBPiece.clearDetail;
    // ---
    TOBBases.clearDetail; TOBBases_O.clearDetail;
    TOBEches.clearDetail; TOBEches_O.clearDetail;
    TOBPorcs.clearDetail; TOBPorcs_O.clearDetail;
    // Fiches
    TOBTiers.clearDetail;
    TOBArticles.clearDetail;
    TOBConds.clearDetail;
    TOBTarif.clearDetail;
    TOBComms.clearDetail;
    TOBCatalogu.clearDetail;
    // Adresses
    TOBAdresses.clearDetail; TOBAdresses_O.clearDetail;
    // Divers
    TOBCXV.clearDetail;
    TOBNomenclature.clearDetail; TOBN_O.clearDetail;
    TOBDim.clearDetail;
    TOBDesLots.clearDetail; TOBLOT_O.clearDetail;
    TOBSerie.clearDetail; TOBSerie_O.clearDetail; TOBSerRel.clearDetail;
    TOBAcomptes.clearDetail; TOBAcomptes_O.clearDetail;
    TOBDispoContreM.clearDetail;
    // Affaires
    TOBAffaire.clearDetail;
    // Comptabilité
    TOBCPTA.clearDetail;
    TOBANAP.clearDetail;
    //Saisie Code Barres
    TOBGSA.clearDetail;
    // MODIF BTP
    // Ouvrages
    TOBOuvrage.clearDetail; TOBOuvrage_O.clearDetail;
    // textes debut et fin
    TOBLIENOLE.clearDetail; TOBLIENOLE_O.clearDetail;
    // retenues de garantie
    TOBPieceRG.clearDetail; TOBPieceRG_O.clearDetail;
    // Bases de tva sur RG
    TOBBasesRG.clearDetail; TOBBASESRG_O.clearDetail;
    // --
    TOBLIGNERG.clearDetail;
    TobLigneTarif.clearDetail; TobLigneTarif_O.clearDetail;
  end;

begin
  //Pour les quantités du bon de préparation
  TOBP_Origine := TOB.Create('', nil, -1);
  TOBP_Origine.Dupliquer(TOBPiece,True,True);
  //----------------------------------------------
  //----- Entrée sur dépôt de destination---------
  //----------------------------------------------
  ValideNumeroPiece;

  Tiers := RechercheTiersDepot(string(TOBPiece.GetValue('GP_DEPOTDEST')));

  //Mise à jour PIECE
  TobPiece.PutValue('GP_DEPOT', TOBPiece.GetValue('GP_DEPOTDEST'));
  TobPiece.PutValue('GP_TIERS', Tiers);           //Tiers correspondant au dépôt
  Depot := TobPiece.GetValue('GP_DEPOT');

  if GP_Tiers.Text <> Tiers then
  begin
    GP_Tiers.Text := Tiers;
    ChargeTiers;
    TobPiece.P('GP_DEPOT', Depot);  //car ChargeTiers, IncidenceTiers (TOBPiece.GetEcran(Self)) change le depôt en fonction de l'écran ;
  end;

  //Mise à jour de LIGNE
  for i:=0 to TobPiece.detail.count-1 do
  begin
    PieceVersLigne(TobPiece, TobPiece.Detail[i]);
  end;
  Litige := True;
  ValideLaPiece;

  //----------------------------------------------
  //----- Entrée sur dépôt de litige -------------
  //----------------------------------------------
  //Gestion du litige => écart constaté entre Qté TRV et Qté TRE
  if (V_PGI.IoError=oeOk) and (TobLitige.detail.count > 0) then
  begin
    //Suppression des lignes sans écart constaté entre Qté TRV et Qté TRE
    for i:=TobPiece.detail.count-1 downto 0 do
    begin
      if TobLitige.FindFirst(['INDICE'],[i],False) = nil then
      begin
        TobPiece_O.detail[i].free;
        TobPiece.detail[i].free;
      end;
    end;

    //Gestion de la pièce TRE sur Dépôt litige
    TransfoPiece    := False;
    PasBouclerCreat := True;
    Action          := taCreat;
    Litige          := False;
    PiecePrecedente := EncodeRefPiece(TobPiece_O);

    //Mise à jour PIECE
    TobPiece.PutValue('GP_DEPOT'    , DepotLitige);
    TobPiece.PutValue('GP_DEPOTDEST', DepotDestination);    //Gestion des dépôts dans LoadLesArticles

    TobArticles.ClearDetail;
    LoadLesArticles;
    TobPiece.PutValue('GP_DEPOTDEST'        , '');
    LoadLesAdresses(TobPiece, TobAdresses);

    Tiers := RechercheTiersDepot(string(TOBPiece.GetValue('GP_DEPOTDEST')));

    //Mise à jour PIECE
    TobPiece.PutValue('GP_TIERS'        , Tiers);           //Tiers correspondant au dépôt
    TobPiece.PutValue('GP_DEPOTDEST'    , '');
    TobPiece.PutValue('GP_DEVENIRPIECE' , '');

    Depot := TobPiece.GetValue('GP_DEPOT');
    if GP_Tiers.Text <> Tiers then
    begin
      GP_Tiers.Text := Tiers;
      ChargeTiers;
      TobPiece.P('GP_DEPOT', Depot);  //car ChargeTiers, IncidenceTiers (TOBPiece.GetEcran(Self)) change le depôt en fonction de l'écran ;
    end;

    ValideNumeroPiece;
    //Mise à jour de LIGNE
    for i:=0 to TobPiece.detail.count-1 do
    begin
      PieceVersLigne(TobPiece, TobPiece.Detail[i]);
      LignePrecedente := Copy(PiecePrecedente, 1, Length(PiecePrecedente) - 1)
        + IntToStr(TobPiece.detail[i].GetValue('GL_NUMORDRE')) + ';';
      TobPiece.Detail[i].PutValue('GL_DEPOT', DepotLitige);
      TobPiece.Detail[i].PutValue('GL_PIECEPRECEDENTE', LignePrecedente);
      TobPiece.Detail[i].PutValue('GL_ETABLISSEMENT', TobPiece_O.Detail[i].GetValue('GL_ETABLISSEMENT'));
      TobPiece.Detail[i].PutValue('GL_QTESTOCK', TobPiece_O.Detail[i].GetValue('GL_QTESTOCK') - TobPiece.Detail[i].GetValue('GL_QTESTOCK'));
      TobPiece.Detail[i].PutValue('GL_QTEFACT' , TobPiece_O.Detail[i].GetValue('GL_QTEFACT') - TobPiece.Detail[i].GetValue('GL_QTEFACT'));
      TobPiece.Detail[i].PutValue('GL_QTERESTE', TobPiece.Detail[i].GetValue('GL_QTESTOCK'));
    end;
    PutValueDetail(TobPiece,'GP_RECALCULER','X');
    CalculeLaSaisie(-1,-1,True);
    CalculFacture(TobPiece, TobBases, TobTiers, TobArticles, TobPorcs, nil, nil, Dev);
    ValideLaPiece;
  end;

  //----------------------------------------------
  //----- Bon de préparation ---------------------
  //----------------------------------------------
  //Gestion du bon de préparation (PRE) en automatique si la réception du transfert
  //correspond à une commande de vente
  if ( V_PGI.IoError=oeOk                                                ) and
     ( copy(Tobpiece.Getvalue('GP_REFEXTERNE'),1,3)='VEN'                ) and
     ( valeurI(copy(Tobpiece.Getvalue('GP_REFEXTERNE'),4,8)) > 0         ) and
     ( valeurI(copy(Tobpiece.Getvalue('GP_REFEXTERNE'),4,8)) < 100000000 ) then
  begin
    RazTobPourPRE;

    Q:=OpenSQL('SELECT GP_DATEPIECE, GP_NUMERO, GP_INDICEG FROM PIECE '
              +' WHERE GP_VIVANTE="X" '
              +' AND GP_NATUREPIECEG="CC" '
              +' AND GP_SOUCHE="GCC" '
              +' AND GP_NUMERO="'+copy(Tobpiece.Getvalue('GP_REFEXTERNE'),4,8)+'"', True);

    St := '';
    if not Q.eof then
      St := 'CC;'+DatetoStr(Q.findfield('GP_DATEPIECE').AsDateTime)+';GCC;'+IntToStr(Q.findfield('GP_NUMERO').AsInteger)+';'+IntToStr(Q.findfield('GP_INDICEG').AsInteger)+';';

    Ferme(Q);

    if St <> '' then
    begin
      StringToCleDoc(St, CleDoc);
      CleDoc.Numligne := 0; CleDoc.NoPersp := 0;

      Action := taModif;
      NewNature := 'PRE';
      TransfoPiece := True;
      DuplicPiece := False;
      GppReliquat := (GetInfoParPiece('PRE', 'GPP_RELIQUAT') = 'X'); { NEWPIECE }

      ChargeLaPiece(False, False);
      GereVivante;
      EtudieReliquat;
      AppliqueTransfoDuplic;
      BloquePiece(Action,True);

      TobQteReste := Tob.Create('_TEMP_', nil, -1);
      try
        //TobBP_Origine => 'TRE'
        //TobPiece_O    => 'CC'
        //TobPiece      => 'PRE'
        for i:=0 to TOBPiece.detail.count-1 do
        begin
          if TobPiece_O.detail[i].G('GL_QTERESTE') > 0 then
          begin
            TobF := TOBP_Origine.FindFirst(['GL_ARTICLE'],[TOBPiece.Detail[i].GetValue('GL_ARTICLE')],False);
            if (TobF <> nil) then
            begin
              //Clé unique de PROPTRANSFLIG = GTL_CODEPTRF,GTL_ARTICLE,GTL_DEPOTDEST
              //donc on peut avoir une seule ligne de transfert pour plusieurs lignes dans la commande
              //d'origine donc on affecte tant qu'il y a du reste à livrer sur les lignes
              T1 := TobQteReste.FindFirst(['ARTICLE'],[TOBPiece.Detail[i].GetValue('GL_ARTICLE')],True);
              if T1 <> nil then
              begin
                nQteTotalStock := T1.G('QTERESTESTOCK');
                nQteTotalFact  := T1.G('QTERESTEFACT');
                nQteStock   := Min(nQteTotalStock, TOBPiece_O.Detail[i].G('GL_QTERESTE'));
                nQteFact    := Min(nQteTotalFact , TOBPiece_O.Detail[i].G('GL_QTERESTE'));
                T1.P('QTERESTESTOCK', nQteTotalStock - nQteStock);
                T1.P('QTERESTEFACT' , nQteTotalFact  - nQteFact);
              end
              else
              begin
                nQteTotalStock := TobF.G('GL_QTERESTE');
                nQteTotalFact  := TobF.G('GL_QTERESTE');
                nQteStock   := Min(nQteTotalStock, TOBPiece_O.Detail[i].G('GL_QTERESTE'));
                nQteFact    := Min(nQteTotalFact , TOBPiece_O.Detail[i].G('GL_QTERESTE'));
                T2 := Tob.Create('_TEMP_', TobQteReste, -1);
                T2.AddChampSupValeur('ARTICLE'      ,TobF.G('GL_ARTICLE') );
                T2.AddChampSupValeur('QTERESTESTOCK',nQteTotalStock - nQteStock);
                T2.AddChampSupValeur('QTERESTEFACT' ,nQteTotalFact  - nQteFact);
              end;

              TobPiece.Detail[i].PutValue('GL_QTESTOCK',nQteStock);
              TobPiece.Detail[i].PutValue('GL_QTEFACT' ,nQteFact );
              TobPiece.Detail[i].PutValue('GL_QTERESTE',nQteStock);
            end
            else
            begin
              TobPiece.Detail[i].PutValue('GL_QTESTOCK',0);
              TobPiece.Detail[i].PutValue('GL_QTEFACT' ,0);
              TobPiece.Detail[i].PutValue('GL_QTERESTE',0);
            end;
          end;
        end;
      finally
        TobQteReste.free;
      end;

      ResGC := GCPieceCorrecte(TOBPiece,TOBArticles,TOBCatalogu);
      if ResGC > 0 then
      begin
        HErr.Execute(ResGC,Caption,'');
        Exit;
      end;
      Ok := GereEcheancesGC(TOBPiece, TOBTiers, TOBEches, TOBAcomptes, TOBPIECERG, Action, DEV, False);
      if not Ok then Exit;

      ValideNumeroPiece;
      ValideLaPiece;
    end;

  end;
  TOBP_Origine.free;

end;

function TFTransfert.CreationTRV : boolean;
var Q : TQuery;
    GereSurSiteDistant : Boolean;
    DepotTransit       : string; 
begin
result:=False; GereSurSiteDistant:=False;
Q:=OpenSQL('Select GDE_SURSITEDISTANT from DEPOTS WHERE GDE_DEPOT="'+TOBPiece.GetValue('GP_DEPOTDEST')+'"',True) ;
if Not Q.EOF then GereSurSiteDistant:=Q.FindField('GDE_SURSITEDISTANT').AsString='X';
Ferme(Q);
if GereSurSiteDistant and GetParamSoc('SO_GCTRV') then Result:=True;
if (not Result) and (GetParamSoc('SO_GCTRV')) then
begin
  DepotTransit := '';
  Q := OpenSQL('Select GDE_DEPOTTRANSIT from DEPOTS WHERE GDE_DEPOT="'+TOBPiece.GetValue('GP_DEPOT')+'"',True);
  if not Q.EOF then DepotTransit := Q.FindField('GDE_DEPOTTRANSIT').AsString ; // = TOBPiece.GetValue('GP_DEPOTDEST') ?
  Ferme(Q);
  if (DepotTransit <> '') then Result := True;
end;
end;

procedure TFTransfert.GP_ETABLISSEMENT_Change(Sender: TObject);
Var Etab,ListeDepot,DepEmett : String ;
begin
  Etab:=GP_ETABLISSEMENT_.Value;
  if Etab='' then Exit ;
  // Appel de la fonction renvoyant la liste des dépôts liés à l'établissement
  ListeDepot:=ListeDepotParEtablissement(Etab);
  if ListeDepot<>'' then
    begin
    DepEmett := GP_DEPOT_.Value;
    GP_DEPOT_.Plus:= ' GDE_SURSITE="X" AND GDE_DEPOT in (' + ListeDepot + ')';
    GP_DEPOT_.Value := DepEmett;
    end
  else
    begin
    DepotChange := True; GP_DEPOTDESTChange(Nil);
    DepotChange := True; GP_DEPOT_Change(Nil);
    end;
end;

procedure TFTransfert.GP_DEPOTDESTChange(Sender: TObject);
Var DepotOri,ListeDepot,StPlus : string;
begin
  if DepotChange then
    begin
    DepotChange := False;
    DepotOri := GP_DEPOT_.Value;
    StPlus := '';
    if (VH_GC.GCMultiDepots) then
       begin
       ListeDepot:=ListeDepotParEtablissement(GP_ETABLISSEMENT_.Value);
       if ListeDepot<>'' then StPlus := ' GDE_DEPOT in (' + ListeDepot + ') AND';
       end;
    if (CtxMode in V_PGI.PGIContexte) and (Action=taCreat) then
      GP_DEPOT_.Plus := StPlus + ' GDE_DEPOT<>"'+ GP_DEPOTDEST.Value + '" AND GDE_SURSITE="X"'
    else
      GP_DEPOT_.Plus := StPlus + ' GDE_DEPOT<>"'+ GP_DEPOTDEST.Value + '"';
    GP_DEPOT_.Value := DepotOri;
    end
  else DepotChange := True;
end;

procedure TFTransfert.GP_DEPOT_Change(Sender: TObject);
Var DepotDest : string;
begin
  if DepotChange then
    begin
    DepotChange := False;
    DepotDest := GP_DEPOTDEST.Value;
    GP_DEPOTDEST.Plus := 'GDE_DEPOT<>"'+ GP_DEPOT_.Value + '"';
    GP_DEPOTDEST.Value := DepotDest;
    end
  else DepotChange := True;
end;

procedure TFTransfert.FormShow(Sender: TObject);
begin
  inherited;
if Action<>TaCreat then PEnteteFO.enabled:=False;
ControlsVisible(PPied,False,['INFOSLIGNE','PCUMUL']);
if VH_GC.GCMultiDepots then
   begin
   HGP_DEPOT_.caption := 'Dépôt émetteur';
   HGP_DEPOTDEST_.caption := 'Dépôt récepteur';
   GP_ETABLISSEMENT_.Visible := True; HGP_ETABLISSEMENT_.Visible := True;
   end
else
   begin
   GP_ETABLISSEMENT_.Visible := False; HGP_ETABLISSEMENT_.Visible := False;
   end;
if SG_px > 0 then GS.ColLengths[SG_px]:=-1;
if ctxFO in V_PGI.PGIContexte then CacherMontant(True);
MAJPanelCumul;
GP_NUMEROPIECE_.Caption:=GP_NUMEROPIECE.Caption;

//Augmente la taille de la police des 3 combos (8 => 11) pour l'utilisation tactile
if CtxFO in V_PGI.PGIContexte then
  begin
  GP_DEPOT_.Font.Size := 11;
  GP_DEPOTDEST.Font.Size := 11;
  GP_ETABLISSEMENT_.Font.Size := 11;
  end;

end;

procedure TFTransfert.CacherMontant(Cacher:boolean);
begin
//Cache les colonnes du grid correspondant à des montants
if Cacher then
  begin
  if SG_Montant>0 then GS.ColWidths[SG_Montant]:=-1;
  if SG_Px>0 then GS.ColWidths[SG_Px]:=-1;
  HMTrad.ResizeGridColumns(GS);
  end;
HGP_TOTALHTDEV_.Visible:=Not Cacher;
LGP_TOTALHTDEV_.Visible:=Not Cacher;
//SG_Total
end;


procedure TFTransfert.GSCellExit(Sender: TObject; var ACol, ARow: Integer;
  var Cancel: Boolean);
begin
  inherited;
if (ACol=SG_QF) or (ACol=SG_QS) or (ACol=SG_Px) then MAJPanelCumul;
end;

procedure TFTransfert.PourlaligneClick(Sender: TObject);
begin
  inherited;
  MAJPanelCumul;
end;

procedure TFTransfert.PourledocumentClick(Sender: TObject);
begin
  inherited;
  MAJPanelCumul;
end;

procedure TFTransfert.DetruitLaPiece ;
var NatureSuivante : string;
begin
inherited;
if CreationTRV then NatureSuivante:='TRV' else NatureSuivante:='TRE';
if ConstruireTREouTRV(NatureSuivante,True) then inherited;
if V_PGI.IoError=oeOk then MAJPanelCumul;
end;

procedure TFTransfert.MAJCaisse;
var Ind : integer;
    TOBL : TOB;
begin
TOBPiece.PutValue('GP_CAISSE', VH_GC.TOBPcaisse.GetValue('GPK_CAISSE')) ;
TOBPiece.PutValue('GP_NUMZCAISSE', VH_GC.TOBPCaisse.GetValue('GPK_NUMZCAISSE')) ;
for Ind:=0 to TOBPiece.Detail.Count-1 do
  begin
  TOBL:=TOBPiece.Detail[Ind] ;
  TOBL.PutValue('GL_CAISSE', VH_GC.TOBPCaisse.GetValue('GPK_CAISSE')) ;
  end;
end;

procedure TFTransfert.ClickDel( ARow : integer ; AvecC,FromUser : boolean; SupDim: boolean=False; TraiteDim: boolean=False ) ;
begin
inherited;
MAJPanelCumul;
end;

procedure TFTransfert.AppliqueTransfoDuplic ;
begin
inherited;
GP_DATEPIECE_.Text:=DateToStr(CleDoc.DatePiece);
end;

Procedure AGLCreerTID ( Parms : array of variant ; nb : integer ) ;
BEGIN
//CreerTransfert(String(Parms[0])) ;
CreerTransfert(String(Parms[0]), String(Parms[1])); 
END ;

Initialization
RegisterAglProc('CreerTID',False,1,AGLCreerTID) ;
end.
