{***********UNITE*************************************************
Auteur  ...... : JS
Créé le ...... : 07/03/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GCVERIFLIGNESERIE ()
Mots clefs ... : TOF;VERIFLIGNESERIE
*****************************************************************}
Unit VerifLigneSerie_TOF ;

Interface

Uses StdCtrls,
     Controls,Windows,
     Classes,M3FP,Buttons,
{$IFDEF EAGLCLIENT}
     eFiche,MaineAGL,
{$ELSE}
     db,FE_Main,
     dbtables,Fiche,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,HMsgBox,
     ParamDBG,Vierge,LicUtil,Ent1,Menus,UTofOptionEdit,TarifUtil,HDimension,LookUp,AGLInitGC,
     UTOF,UTOB,EntGC,HTB97,Facture,FactUtil,TiersUtil,HPanel,HSysMenu,UTofConsultStock;

function GCLanceFiche_VerifLigneSerie(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_VERIFLIGNESERIE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

  private
    GTRA : THGrid;
    HMTrad: THSystemMenu;
    BParamListe,BImprimer : TToolbarButton97;
    BZoomArticle,BZoomClient,BZoomFour,BZoomPiece : TBitBtn;
    PEtat : RPARETAT;
    TobLigneSerie : TOB;
    IdentArticle, OldArticle, NatureAuxi : string;
    ListePiece,stTitresLigne : string;

    /// Boutons
    procedure BZoomArticleClick(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure BZoomClientClick(Sender: TObject);
    procedure BZoomFourClick(Sender: TObject);
    procedure ParamListeClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    /// Grid
    procedure GDblClick(Sender: TObject);
    procedure InitialiseForm;
    procedure EtudieColsListes ;
    procedure AppelLaPiece;
    procedure AddChampListe(var NomListe,NomChamp : string);
    ///Tob
    procedure ChargeTobMouvement;
    ///Fiche
    function  RecupCriteres(stPrefixe : string) : string;
    procedure ArticleElipsisClick(Sender : TObject);
    procedure ArticleExit(Sender : TObject);
    procedure RechercheDimensions;
    function  RechercheLibelleDim(TobArt : TOB) : string;
    ///Impression
    function  PrepareImpression : integer;

  end ;

const
// libellés des messages
TexteMessage: array[1..2] of string = (
          {1}  'Vous devez renseigner un article'
          {2} ,'Vous devez renseigner un tiers'
              );

Implementation

function GCLanceFiche_VerifLigneSerie(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:='';
if Nat='' then exit;
if Cod='' then exit;
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

{==============================================================================================}
{================================== Procédure de la TOF =======================================}
{==============================================================================================}
procedure TOF_VERIFLIGNESERIE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_VERIFLIGNESERIE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_VERIFLIGNESERIE.OnUpdate ;
begin
  Inherited ;
GTRA.VidePile(false) ;
TobLigneSerie.ClearDetail ;
if THEdit(GetControl('CODEARTICLE')).Focused then ArticleExit(GetControl('CODEARTICLE'));
if GetControlText('CODEARTICLE')='' then
   begin
   SetFocusControl('CODEARTICLE');
   HShowMessage('0;'+Ecran.Caption+';'+TexteMessage[1]+';W;O;O;O;','','');
   exit;
   end;
if GetControlText('TIERS')='' then
   begin
   SetFocusControl('TIERS');
   HShowMessage('0;'+Ecran.Caption+';'+TexteMessage[2]+';W;O;O;O;','','');
   exit;
   end;
ChargeTobMouvement;
TobLigneSerie.PutGridDetail(GTRA,false,false,stTitresLigne);
THPanel(GetControl('PCUMULTRA')).Caption := 'Totaux  ('+IntToStr(TobLigneSerie.Detail.Count)+' lignes)';
end ;


procedure TOF_VERIFLIGNESERIE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_VERIFLIGNESERIE.OnArgument (S : String ) ;
begin
  Inherited ;
NatureAuxi := Trim(S);
if NatureAuxi = '' then NatureAuxi := 'CLI';
//***Récup des tables et des préfixes
ListePiece := 'GL_NATUREPIECEG;GL_NUMERO;GL_NUMLIGNE;GL_SOUCHE;GL_INDICEG;GL_VIVANTE;'
             +'GL_ARTICLE;GL_DEPOT;GL_DATEPIECE;GL_TIERS;GLS_IDSERIE';
/////////Affectations
GTRA:=THGrid(GetControl('GTRA'));
GTRA.OnDblClick := GDblClick;
THEdit(GetControl('CODEARTICLE')).OnElipsisClick := ArticleElipsisClick;
THEdit(GetControl('CODEARTICLE')).OnDblClick := ArticleElipsisClick;
THEdit(GetControl('CODEARTICLE')).OnExit := ArticleExit;
//Boutons
BParamListe := TToolbarButton97(GetControl('BParamListe'));
BImprimer := TToolbarButton97(GetControl('BImprimer'));
BZoomArticle := TBitBtn(GetControl('BZOOMARTICLE'));
BZoomClient := TBitBtn(GetControl('BZOOMCLIENT'));
BZoomFour := TBitBtn(GetControl('BZOOMFOUR'));
BZoomPiece := TBitBtn(GetControl('BZOOMPIECE'));
BParamListe.OnClick := ParamListeClick;
BImprimer.OnClick := BImprimerClick;
bZoomArticle.OnClick := BZoomArticleClick;
BZoomClient.OnClick := BZoomClientClick;
BZoomFour.OnClick := BZoomFourClick;
BZoomPiece.OnClick := BZoomPieceClick;
////////Initialisation
BParamListe.Visible := (V_PGI.PassWord = CryptageSt(DayPass(Date)));
//Grids
GTRA.ListeParam := 'GCVERIFLIGNESERIE';
EtudieColsListes ;
InitialiseForm;
HMTrad.ResizeGridColumns(GTRA);
IdentArticle := '' ; OldArticle := '';
//Tob
TobLigneSerie := TOB.Create('',nil,-1) ;
//Etat
PEtat.Tip:='E'; PEtat.Nat:='GZJ'; PEtat.Modele:='VER'; PEtat.Titre:=Ecran.Caption;
PEtat.Apercu:=True; PEtat.DeuxPages:=False; PEtat.First:=True;
PEtat.stSQL := 'GZJ_UTILISATEUR = "'+ V_PGI.USer +'"';
//Menu zoom
PopZoom97(TToolbarButton97(GetControl('BZOOM')),TPopupMenu(GetControl('POPMENU'))) ;
end ;

procedure TOF_VERIFLIGNESERIE.OnClose ;
begin
  Inherited ;
TobLigneSerie.Free ;
end ;

////////////////////////////////////////////////////////////////////////////////
//Grid
procedure TOF_VERIFLIGNESERIE.GDblClick(Sender: TObject);
begin
if GTRA.Row > 0 then AppelLaPiece;
end;

////////////////////////////////////////////////////////////////////////////////
//Tob
procedure TOF_VERIFLIGNESERIE.ChargeTobMouvement;
var TobD : TOB;
    i_ind : integer;
    stSQL : string;
begin
stSQL := 'SELECT ' + StringReplace(ListePiece,';',',',[rfReplaceAll])
  + ' FROM LIGNE LEFT JOIN LIGNESERIE ON '
  + ' GL_NATUREPIECEG=GLS_NATUREPIECEG AND GL_NUMLIGNE=GLS_NUMLIGNE'
  + ' AND GL_SOUCHE=GLS_SOUCHE AND GL_INDICEG=GLS_INDICEG'
  + ' AND GL_NUMERO=GLS_NUMERO ' + RecupCriteres('GL') + ' AND GL_NUMEROSERIE="X" AND '
  + RecupWhereNaturePiece('PHY','GL');
TobLigneSerie.LoadDetailFromSQL(stSQL,True);
if TobLigneSerie.Detail.Count = 0 then exit;
TobLigneSerie.Detail[0].AddChampSup('(NATUREPIECE)',True);
TobLigneSerie.Detail[0].AddChampSup('(DEPOT)',True);
TobLigneSerie.Detail[0].AddChampSup('(QUANTITE)',True);
//suppression des pièces non vivantes si elles ont été transformées
for i_ind := TobLigneSerie.Detail.Count -1 downto 0 do
    begin
    TobD := TobLigneSerie.Detail[i_ind];
    if (TobD.GetValue('GL_VIVANTE') = '-') and
       (GetInfoParPiece(TobD.GetValue('GL_NATUREPIECEG'),'GPP_ACTIONFINI') = 'TRA') then TobD.Free
    else begin
         TobD.PutValue('(NATUREPIECE)',RechDom('GCNATUREPIECEG',TobD.GetValue('GL_NATUREPIECEG'),false));
         TobD.PutValue('(DEPOT)',RechDom('GCDEPOT',TobD.GetValue('GL_DEPOT'),false));
         end;
    end;
end;

////////////////////////////////////////////////////////////////////////////////
//Grid
procedure TOF_VERIFLIGNESERIE.EtudieColsListes ;
Var NomCol,LesCols : String ;
    ichamp : integer ;
begin
stTitresLigne:=GTRA.Titres[0] ; 
LesCols := stTitresLigne;
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        begin
        ichamp:=ChampToNum(NomCol) ;
        if ichamp>=0 then AddChampListe(ListePiece,NomCol);
        end;
Until ((LesCols='') or (NomCol=''));
end ;

procedure TOF_VERIFLIGNESERIE.InitialiseForm;
var TitreEcran : string;
begin
//TitreEcran := 'Traçabilité des numéros de série';
if NatureAuxi = 'CLI' then
   begin
   SetControlProperty('TIERS','DataType','GCTIERSCLI');
   TitreEcran := TraduireMemoire('Vérification vente série');
   SetControlEnabled('BZOOMCLIENT',True);
   SetControlEnabled('BZOOMFOUR',False);
   SetControlCaption('TTIERS',TraduireMemoire('Client'));
   end else
   begin
   SetControlProperty('TIERS','DataType','GCTIERSFOURN');
   TitreEcran := TraduireMemoire('Vérification achat série');
   SetControlEnabled('BZOOMCLIENT',False);
   SetControlEnabled('BZOOMFOUR',True);
   SetControlCaption('TTIERS',TraduireMemoire('Fournisseur'));
   end;
Ecran.Caption := TitreEcran;
UpdateCaption(Ecran);
SetControlProperty('CODEARTICLE','Plus','AND GA_NUMEROSERIE="X"');
SetControlText('TLIBELLEDIM','');
end;

////////////////////////////////////////////////////////////////////////////////
//Fiche
function TOF_VERIFLIGNESERIE.RecupCriteres(stPrefixe : string) : string;
var stPrefRef : string;
begin
if stPrefixe = 'GL' then stPrefRef := 'GLS' else stPrefRef := stPrefixe;
result := ' WHERE ' + stPrefixe + '_ARTICLE="' + IdentArticle + '"';
if GetControlText('TIERS') <> '' then
   result := result + ' AND ' + stPrefixe + '_TIERS="' + GetControlText('TIERS') + '"';
if GetControlText('DEPOT') <> '' then
   result := result + ' AND ' + stPrefixe + '_DEPOT="' + GetControlText('DEPOT') + '"';
if GetControlText('IDSERIE') <> '' then
   result := result + ' AND ' + stPrefRef + '_IDSERIE>="' + GetControlText('IDSERIE') + '"';
if GetControlText('IDSERIE_') <> '' then
   result := result + ' AND ' + stPrefRef + '_IDSERIE<="' + GetControlText('IDSERIE_') + '"';
result := result + ' AND (' + stPrefixe + '_DATEPIECE>="' + USDateTime(strToDate(GetControlText('DATEPIECE'))) + '"'
  + ' AND ' + stPrefixe + '_DATEPIECE<="' + USDateTime(strToDate(GetControlText('DATEPIECE_'))) + '")';
result := result + ' AND (' + stPrefixe + '_DATELIVRAISON>="' + USDateTime(strToDate(GetControlText('DATELIVRAISON'))) + '"'
  + ' AND ' + stPrefixe + '_DATELIVRAISON<="' + USDateTime(strToDate(GetControlText('DATELIVRAISON_'))) + '")';
end;

Procedure TOF_VERIFLIGNESERIE.ArticleElipsisClick(Sender : TObject);
var CodeArt, StatutArt : string;
begin
CodeArt := GetControlText('CODEARTICLE');
StatutArt := '' ;  IdentArticle := '' ;
DispatchRecherche(THCritMaskEdit(GetControl('CODEARTICLE')), 1,'', 'GA_CODEARTICLE=' + CodeArt + ';RETOUR_CODEARTICLE=X', '');
RechercheDimensions;
OldArticle := GetControlText('CODEARTICLE');
end;

Procedure TOF_VERIFLIGNESERIE.ArticleExit(Sender : TObject);
begin
if GetControlText('CODEARTICLE') = OldArticle then exit;
RechercheDimensions;
OldArticle := GetControlText('CODEARTICLE');
end;

procedure TOF_VERIFLIGNESERIE.RechercheDimensions;
var CodeArt, StatutArt : string;
    QArt : TQuery;
    TobArt : TOB;
begin
CodeArt := GetControlText('CODEARTICLE');  IdentArticle := CodeArt ;
QArt :=  OpenSQL('SELECT GA_ARTICLE,GA_STATUTART FROM ARTICLE WHERE GA_CODEARTICLE="' + CodeArt
     + '" AND GA_STATUTART<>"DIM"',True);
if not QArt.Eof then
   begin
   IdentArticle := QArt.FindField('GA_ARTICLE').AsString;
   StatutArt := QArt.FindField('GA_STATUTART').AsString;
   end;
Ferme(QArt);
if StatutArt = 'GEN' then
   begin
   TobArt := TOB.Create('',nil,-1);
   if ChoisirDimension(IdentArticle,TobArt) then
      begin
      SetControlText('TLIBELLEDIM',RechercheLibelleDim(TobArt));
      IdentArticle := TobArt.GetValue('GA_ARTICLE');
      end;
   TobArt.Free;
   end else SetControlText('TLIBELLEDIM','');
end;

function TOF_VERIFLIGNESERIE.RechercheLibelleDim(TobArt : TOB) : string;
var i_indDim : integer;
    GrilleDim,CodeDim,LibDim,StDim : string;
begin
StDim:='';
for i_indDim := 1 to MaxDimension do
    begin
    GrilleDim := TOBArt.GetValue ('GA_GRILLEDIM' + IntToStr (i_indDim));
    CodeDim := TOBArt.GetValue ('GA_CODEDIM' + IntToStr (i_indDim));
    LibDim := GCGetCodeDim (GrilleDim, CodeDim, i_indDim);
    if LibDim <> '' then
       if StDim='' then StDim:=StDim + LibDim
       else StDim := StDim + ' - ' + LibDim;
    end;
Result := stDim;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_VERIFLIGNESERIE.AddChampListe(var NomListe,NomChamp : string);
var TokenListe,st : string;
begin
TokenListe := NomListe;
Repeat
    st := Trim(ReadTokenSt(TokenListe));
    if st = NomChamp then exit;
Until (TokenListe = '') or (st = '');
NomListe := NomListe + ';' + NomChamp;
end;

function TOF_VERIFLIGNESERIE.PrepareImpression : integer;
var TobImp,TOBI, TOBD : TOB;
    i_ind : integer;
begin
//////Nom table à renseigner
result := 0;
ExecuteSQL('DELETE FROM GCTMPTRACABILITE WHERE ' + PEtat.stSQL);
TobImp := TOB.Create('',nil,-1);
for i_ind := 0  to TobLigneSerie.Detail.Count -1 do
    begin
    TOBD := TobLigneSerie.Detail[i_ind];
    TOBI := TOB.Create('GCTMPTRACABILITE',TobImp,-1);
    TOBI.PutValue('GZJ_UTILISATEUR',V_PGI.USer);
    TOBI.PutValue('GZJ_COMPTEUR',i_ind+1);
    TOBI.PutValue('GZJ_LOTSERIE',TOBD.GetValue('GLS_IDSERIE'));
    TOBI.PutValue('GZJ_DEPOT',TOBD.GetValue('GL_DEPOT'));
    TOBI.PutValue('GZJ_DATE',TOBD.GetValue('GL_DATEPIECE'));
    TOBI.PutValue('GZJ_DATELIVRAISON',TOBD.GetValue('GL_DATELIVRAISON'));
    TOBI.PutValue('GZJ_NATUREPIECEG',TOBD.GetValue('GL_NATUREPIECEG'));
    TOBI.PutValue('GZJ_NUMERO',TOBD.GetValue('GL_NUMERO'));
    end;
TobImp.InsertOrUpdateDB(false);
TobImp.Free;
end;

procedure TOF_VERIFLIGNESERIE.AppelLaPiece;
var TOBD : TOB;
    CleDoc : R_CleDoc;
begin
if TobLigneSerie.Detail.Count < GTRA.Row then exit;
TOBD := Tob(GTRA.Objects[0,GTRA.Row]);
if TOBD = nil then exit;
CleDoc.NaturePiece:=TOBD.GetValue('GL_NATUREPIECEG');
CleDoc.Souche:=TOBD.GetValue('GL_SOUCHE');
CleDoc.DatePiece:=TOBD.GetValue('GL_DATEPIECE');
CleDoc.NumeroPiece:=TOBD.GetValue('GL_NUMERO');
CleDoc.Indice:=TOBD.GetValue('GL_INDICEG');
SaisiePiece(CleDoc,taConsult);
end;

////////////////////////////////////////////////////////////////////////////////
/////Boutons

procedure TOF_VERIFLIGNESERIE.BImprimerClick(Sender: TObject);
begin
if TobLigneSerie.Detail.Count = 0 then exit;
SetControlText('LIBELLEARTICLE',RechDom('GCARTICLE',IdentArticle,false) + '  '
                 + GetControlText('TLIBELLEDIM'));
SetControlText('RAISONSOCIALE',RechDom(THEdit(GetControl('TIERS')).DataType,
                TiersAuxiliaire(GetControlText('TIERS')),false));
EntreeOptionEdit(PEtat.Tip,PEtat.Nat,PEtat.Modele,PEtat.Apercu,PEtat.DeuxPages,PEtat.First,
                 TPageControl(GetControl('PAGECONTROL')),PEtat.stSQL,PEtat.Titre,PrepareImpression);
ExecuteSQL('DELETE FROM GCTMPTRACABILITE WHERE ' + PEtat.stSQL);
end;

procedure TOF_VERIFLIGNESERIE.BZoomArticleClick(Sender: TObject);
begin
{$IFNDEF GPAO}
AGLLanceFiche('GC','GCARTICLE','',IdentArticle,'ACTION=CONSULTATION;MONOFICHE') ;
{$ELSE}
V_PGI.DispatchTT(7, taConsult, IdentArticle, 'MONOFICHE', '');
{$ENDIF}
end;

procedure TOF_VERIFLIGNESERIE.BZoomPieceClick(Sender: TObject);
begin
AppelLaPiece;
end;

procedure TOF_VERIFLIGNESERIE.BZoomClientClick(Sender: TObject);
var CodeAuxi : string;
begin
if GetControlText('TIERS') = '' then exit;
CodeAuxi := TiersAuxiliaire(GetControlText('TIERS'));
if CodeAuxi <> '' then AGLLanceFiche('GC','GCTIERS','',CodeAuxi,'ACTION=CONSULTATION;MONOFICHE') ;
end;

procedure TOF_VERIFLIGNESERIE.BZoomFourClick(Sender: TObject);
begin
if GetControlText('TIERS') = '' then exit;
AGLLanceFiche('GC','GCFOURNISSEUR','',TiersAuxiliaire(GetControlText('TIERS')),'ACTION=CONSULTATION;MONOFICHE');
end;

procedure TOF_VERIFLIGNESERIE.ParamListeClick(Sender: TObject);
var Liste : string;
begin
Liste := GTRA.ListeParam;
{$IFDEF EAGLCLIENT}
ParamListe(Liste,nil) ;
{$ELSE}
ParamListe(Liste,nil,nil) ;
{$ENDIF}
EtudieColsListes;
HMTrad.ResizeGridColumns(GTRA);
TobLigneSerie.ClearDetail ;
ChargeTobMouvement;
TobLigneSerie.PutGridDetail(GTRA,false,false,stTitresLigne);
THPanel(GetControl('PCUMULTRA')).Caption := 'Totaux ('+IntToStr(TobLigneSerie.Detail.Count)+' lignes)';
end;

Initialization
  registerclasses ( [ TOF_VERIFLIGNESERIE ] ) ;
end.

