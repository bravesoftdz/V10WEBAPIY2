{***********UNITE*************************************************
Auteur  ...... : Paul Chapuis
Créé le ...... : 18/12/2000
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : GCENCOURS ()
Mots clefs ... : TOF;GCENCOURS
*****************************************************************}
Unit UTOFGCEncours ;

Interface

Uses  StdCtrls, Controls, Classes,  forms, sysutils, ComCtrls,grids ,graphics,windows,vierge,Messages,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOF,TiersUtil,Utob, AGLInit,EntGC,M3FP, FactUtil,Facture,Ent1,ExtCtrls,
     HPanel,SaisUtil,HSysMenu,FactCpta,HTB97,UTofOptionEdit,UtilPGI ,EcrPiece_TOF,Paramsoc,uEntCommun;

Type
  TOF_GCENCOURS = Class (TOF)
  private
    VOR         : string ;       // etat reel du risque client   (vert,orange,rouge = V,O,R)
    TobTiers    : Tob ;
    Action      : TActionFiche ;
    procedure AfficherDetail(quoi : string);
  protected
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnUpdate                 ; override ;
  end ;

  TOF_GCENCOURSDETGC = Class (TOF)
  private
    Action      : TActionFiche ;
    Splits      : Array of TSplitter;
    LesColonnes : Array of string ;
    GS          : THGRID ;
    G           : Array of THGRID;
    TobA        : Array of TOB ;
    TobTiers    : Tob ;
    colDateEch  : integer;
    TypeEcritures   : string;
    Reaffiche : boolean;
    stRisquecpta,stRisqueGC : string ;
    BImprimer: TToolbarButton97;
    procedure G1DessineCell (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
    procedure G2DessineCell (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
    procedure G0DessineCell (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
    procedure G2GetCellCanvas(ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
    procedure GSLigneDClick (Sender: TObject);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    Procedure AddSplitter(Ig :Integer) ;
    Procedure AfficherJustif ;
// cd 050701
    Procedure AfficherEC ;
    Procedure AfficherECDetail (TobP: TOB) ;
// cd 050701
    procedure BImprimerClick(Sender: TObject);
    function PrepareImpression : integer ;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  Protected
    procedure OnArgument (Arguments : String ) ; override ;
    procedure OnLoad                           ; override ;
    procedure OnClose                          ; override ;
  end ;

Const MAXG=2;

Procedure RisqueTiersGCDetail ( TOBTiers,TobPieces : TOB );
Procedure EnCoursJustifSoldeDetail (Aux : String ; DateButoir : tDateTime; TobSolde : TOB; TypeEcr : string);
Procedure EnCoursRegleNonEchuDetail (Aux : String ; DateButoir : tDateTime; TobNonECh : TOB);
Procedure DessineCell (GS : THGRID; ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
Function  EnCoursEchuNonRegle (Aux : String ; DateButoir : tDateTime): double;

Implementation
uses UFonctionsCBP;

procedure TOF_GCENCOURS.OnLoad ;
var RisqueGC,RisqueCPTA,Solde,RegleNonEchu,EchuNonRegle : double;
    v : string;
    RisqueTiers,P1, P2 : double;
    iNbJoursEncours : integer;
begin
Inherited ;
If GetControl('TIERS') <> nil then SetControlText('TIERS',TobTiers.getValue('T_TIERS'));
TobTiers.putEcran(Ecran);
RisqueGC:=RisqueTiersGC(TobTiers);
Solde:=TobTiers.GetValue('T_TOTALDEBIT')-TobTiers.GetValue('T_TOTALCREDIT') ;
//CEGID-CCMX LE 15/09/2006 Debut {Evolution}
iNbJoursEncours := GetParamSocSecur('SO_NBJENCOURS', 0);

RegleNonEchu:=-EncoursRegleNonEchu(TobTiers.GetValue('T_AUXILIAIRE'),V_PGI.DateEntree,iNbJoursEncours) ;
RisqueCPTA:=Solde+RegleNonEchu ;
EchuNonRegle:=EnCoursEchuNonRegle(TobTiers.GetValue('T_AUXILIAIRE'),V_PGI.DateEntree) ;

vor:='V';
RisqueTiers:=RisqueGC+RisqueCPTA;
P1:=TobTiers.GetValue('T_CREDITACCORDE') ;
P2:=TobTiers.GetValue('T_CREDITPLAFOND') ;
if ((P1>0 ) and (RisqueTiers > P1)) then vor:='O' ;
if ((P2>0 ) and (RisqueTiers > P2)) then vor:='R' ;

SetControlText('RISQUEGC',FloatToStrF(RisqueGC,ffNumber,15,2));
SetControlText('RISQUECPTA',FloatToStrF(RisqueCpta,ffNumber,15,2));
SetControlText('RISQUE',FloatToStrF(RisqueTiers ,ffNumber,15,2));
SetControlText('NONECHU',FloatToStrF(RegleNonEchu,ffNumber,15,2));
SetControlText('ECHUNONREGLE',FloatToStrF(EchuNonRegle,ffNumber,15,2));

v:=TobTiers.getValue('T_ETATRISQUE');   // etat risque forcé si <> ''
SetControlProperty('FORCERISQUE','checked',(v<>'')) ;
SetControlText('VOR',vor);
if (v='') then v:=vor;
SetControlProperty('RB_VERT','checked',(V='V')) ;
SetControlProperty('RB_ROUGE','checked',(V='R')) ;
SetControlProperty('RB_ORANGE','checked',(V='O')) ;
end ;

procedure TOF_GCENCOURS.OnArgument (S : String ) ;
var St,Critere:string ;
    x:integer;
    ChampMul,ValMul : string;
begin
Inherited ;
St:=S;
Action:=taModif ;
Repeat
    Critere:=uppercase(ReadTokenSt(S)) ;
    if Critere<>'' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           if ChampMul='ACTION' then
             begin
             if ValMul='CREATION' then BEGIN Action:=taCreat ; END ;
             if ValMul='MODIFICATION' then BEGIN Action:=taModif ; END ;
             if ValMul='CONSULTATION' then BEGIN Action:=taConsult ; END ;
             end
//           else if ChampMul='VOR' then Vor:=ValMul
           ;
           end;
        end;
until  Critere='';
if (Action=taConsult) or (not ExJaiLeDroitConcept(TConcept(gcForceRisque),False)) then
    begin
    FicheReadOnly(Ecran) ;
    setControlenabled('BROUGE',True);
    setControlenabled('BVERT',True);
    setControlenabled('BORANGE',True);
    end;

TobTiers:=LaTob;
LaTob:=Nil;
end ;

procedure TOF_GCENCOURS.OnUpdate ;
begin
  Inherited ;
  TobTiers.PutValue('T_ETATRISQUE',GetControlText('T_ETATRISQUE'));
end ;

procedure TOF_GCENCOURS.OnClose ;
begin
  Inherited ;
end ;

Procedure TOF_GCENCOURS.AfficherDetail (Quoi : string)  ;
begin
inherited ;
TheTob:=TobTiers;
AGLLanceFiche ('GC','GCENCOURSGC','','',ActionToString(Action)+';'+Quoi) ;
TheTob:=nil;
end;

//=== Encours detail ============================================================

procedure TOF_GCENCOURSDETGC.OnArgument (Arguments : String ) ;
var i ,x: integer ;
    nbCol : array[0..MAXG] of integer;
    St,Nam,Critere : string ;
    ig:integer;
    ChampMul,ValMul : string;
    RisqueGC,RisqueCPTA,Solde,RegleNonEchu : double;
    iNbJoursEncours : integer;
begin
inherited ;
TobTiers:=LaTob; LaTob:=Nil;
st:=ARguments ;
stRisquecpta:='';
stRisqueGC:='';
Action:=taModif ;
  //CEGID-CCMX LE 15/09/2006 Debut {Evolution}
  iNbJoursEncours := GetParamSocSecur('SO_NBJENCOURS', 0);

Repeat
  Critere:=Trim(ReadTokenSt(st)) ;
  if Critere<>'' then
    begin
    x:=pos('=',Critere);
    if x<>0 then
       begin
       ChampMul:=copy(Critere,1,x-1) ;
       ValMul:=copy(Critere,x+1,length(Critere)) ;
       if ChampMul='RISQUECPTA' then stRisquecpta:=ValMul ;
       if ChampMul='RISQUEGC' then stRisqueGC:=ValMul ;
       if ChampMul='ACTION' then
           begin
           if ValMul='CREATION' then BEGIN Action:=taCreat ; END ;
           if ValMul='MODIFICATION' then BEGIN Action:=taModif ; END ;
           if ValMul='CONSULTATION' then BEGIN Action:=taConsult ; END ;
           end
       end ;
    end ;
until  Critere='' ;

if stRisquecpta = '' then
   begin
   Solde:=TobTiers.GetValue('T_TOTALDEBIT')-TobTiers.GetValue('T_TOTALCREDIT') ;
   RegleNonEchu:=-EncoursRegleNonEchu(TobTiers.GetValue('T_AUXILIAIRE'),V_PGI.DateEntree,iNbJoursEncours) ;
   RisqueCPTA:=Solde+RegleNonEchu ;
   stRisqueCPTA:=FloatToStrF(RisqueCPTA,ffNumber,15,2)
   end;

if stRisqueGC ='' then
   begin
   RisqueGC:=RisqueTiersGC(TobTiers);
   stRisqueGC:=FloatToStrF(RisqueGC,ffNumber,15,2)
   end;

setcontroltext('RISQUEGC',stRisqueGC);
setcontroltext('RISQUECPTA',stRisquecpta);

Reaffiche:=false;
SetLength(TobA,MAXG+1); SetLength(LesColonnes,MAXG+1); SetLength(G,MAXG+1);SetLength(Splits,MAXG+1);
for ig:=0 to MAXG do TobA[ig]:=tob.create ('_A afficher',Nil,-1 );

//'Détail encours commercial';
NbCol[0]:=6;
LesColonnes[0]:='FIXED;BIDON;GP_NATUREPIECEG;GP_DATEPIECE;GP_NUMERO;GP_TOTALTTC' ;
G[0]:=THGRID(GetControl('G0'));
G[0].PostDrawCell:= G0DessineCell;

//'Détail réglements non échus';
NbCol[1]:=10;
LesColonnes[1]:='FIXED;BIDON;E_DATEECHEANCE;E_DATECOMPTABLE;E_JOURNAL;E_NUMEROPIECE;E_LIBELLE;E_REFINTERNE;E_MODEPAIE;Solde';
G[1]:=THGRID(GetControl('G1'));
G[1].PostDrawCell:= G1DessineCell;

//'Justifiactif du solde comptable';
NbCol[2]:=13;
LesColonnes[2]:='FIXED;BIDON;E_DATECOMPTABLE;E_JOURNAL;E_NUMEROPIECE;E_LIBELLE;E_REFINTERNE;E_DATEECHEANCE;E_MODEPAIE;E_DEBIT;E_CREDIT;soldelettre;E_LETTRAGE';
G[2]:=THGRID(GetControl('G2'));
G[2].PostDrawCell:= G2DessineCell;
G[2].GetCellCanvas:= G2GetCellCanvas;

for ig:=0 to MAXG do
   begin
   G[ig].OnRowEnter:=GSRowEnter ;
   G[ig].OnRowExit:=GSRowExit ;
   G[ig].OnDblClick:=GSLigneDClick ;
   G[ig].ColCount:=NbCol[ig];
   G[ig].ColWidths[0]:=15;
   G[ig].ColWidths[1]:=0; // colonne vide PROVISOIRE pour empecher l'affichage du combo
   St:=LesColonnes[ig] ;
   GS:=G[ig];
    for i:=0 to GS.ColCount-1 do
       BEGIN
       if i>1 then GS.ColWidths[i]:=100;
       Nam:=uppercase(ReadTokenSt(St)) ;
       if (Nam='GP_DATEPIECE')  or (Nam='E_DATECOMPTABLE') then begin GS.ColAligns[i]:=tacenter; GS.ColWidths[i]:=80 end
       else if (Nam='E_DATEECHEANCE')  then begin GS.ColAligns[i]:=tacenter; GS.ColWidths[i]:=80 ; colDateEch:=i; end
       else if (Nam='GP_NATUREPIECEG') or (Nam='E_MODEPAIE') then GS.ColFormats[i]:='CB='+ChampToTT(Nam)
       else if (Nam='GP_TOTALTTC') or (Nam='E_DEBIT') or (Nam='E_CREDIT') or (Nam='SOLDE') or (Nam='SOLDELETTRE')then
            begin GS.ColWidths[i]:=80; GS.ColFormats[i]:='#,##0.00'; GS.ColAligns[i]:=taRightJustify ;  end
       else if (Nam='E_LIBELLE') or (Nam='E_REFINTERNE') then GS.ColWidths[i]:=130
       else if (Nam='E_LETTRAGE') then GS.ColWidths[i]:=40
       else if (Nam='E_JOURNAL') then GS.ColWidths[i]:=40
       else if (Nam='E_NUMEROPIECE') then GS.ColWidths[i]:=60
       ;
       END ;
    GS.options:=GS.Options-[goEditing] ;
    TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS) ;
   end;
BImprimer:=TToolbarButton97(GetControl('Bimprimer')) ;
BImprimer.OnClick:=BImprimerClick;
UpdateCaption(Ecran) ;
if (ecran <> Nil)then Ecran.OnKeyDown:=FormKeyDown ;
end;

procedure TOF_GCENCOURSDETGC.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of


    VK_F10 :  {Contacts}  if (ssAlt in Shift)  then
              begin
              AGLLanceFiche('YY','YYCONTACT','T;'+TobTiers.GetValue('t_AUXILIAIRE'),'', ActionToString(Action)+';TYPE=T;'+'TYPE2='+TobTiers.GetValue('T_NATUREAUXI')+';PART='+TobTiers.GetValue('T_PARTICULIER')+';TITRE='+TobTiers.GetValue('T_LIBELLE'));
              end;
    VK_F7 : {Actions} if (ssAlt in Shift) then
                       begin Key:=0 ; AGLLanceFiche('RT','RTACTIONS_TIERS','RAC_TIERS='+TobTiers.GetValue('T_TIERS'),'',ActionToString(Action)+';NOCHANGEPROSPECT') ;  end;

END;
end;

procedure TOF_GCENCOURSDETGC.G0DessineCell(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
begin
DessineCell(G[0],Acol,Arow,Canvas,AState);
end;
procedure TOF_GCENCOURSDETGC.G1DessineCell(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
begin
DessineCell(G[1],Acol,Arow,Canvas,AState);
end;
procedure TOF_GCENCOURSDETGC.G2DessineCell(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
begin
DessineCell(G[2],Acol,Arow,Canvas,AState);
end;

procedure TOF_GCENCOURSDETGC.G2GetCellCanvas(ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
begin
if (Acol=colDateEch) and (ARow>=G[2].Fixedrows) and (G[2].cells[colDateEch,ARow]<>'') and (strToDate(G[2].cells[colDateEch,Arow])< V_PGI.DateEntree) then Canvas.font.color:=clRed;
end;


Procedure TOF_GCENCOURSDETGC.OnLoad  ;
var ig,nbVisible,Hauteur,i,ii:integer;
begin
inherited ;
nbVisible:=0;
TypeEcritures:=GetControlText('NONLETTRE');
for ig:=0 to  MAXG do
    begin
    case ig of
    0 : if not Reaffiche then  RisqueTiersGCDetail ( TobTiers ,TobA[ig]);
    1 : if not Reaffiche then  EnCoursRegleNonEchuDetail (TobTiers.GetValue('T_AUXILIAIRE'),V_PGI.DateEntree,TobA[ig]) ;
    2 : begin
        EnCoursJustifSoldeDetail (TobTiers.GetValue('T_AUXILIAIRE'),V_PGI.DateEntree, TobA[ig], TypeEcritures);
        end;
    end;
    If TobA[ig].detail.count=0 then
         SetControlVisible( 'P'+intToStr(ig),False)
         else begin
         inc(nbVisible);
         if nbVisible=1 then SetControlProperty('P'+intToStr(Ig),'Alig',alClient);
         SetControlVisible( 'P'+intToStr(ig),True);
         if (not Reaffiche) and (ig < MAXG) then AddSplitter(Ig);
         For i:=0 to MAXG-1 do SetControlvisible('PBASE'+intToStr(i),True);
         G[ig].VidePile(false) ;
         TobA[ig].PutGridDetail(G[ig],True,True,LesColonnes[ig],True);
         if Ig=2 then begin G[ig].cells[11,0]:='Solde lettr.';  G[ig].cells[12,0]:='Lettre'; G[ig].cells[3,0]:='Journal'; end;
         For i:=G[ig].fixedrows to G[ig].RowCount-1 do
             begin
             for ii:= G[ig].fixedcols to G[ig].ColCount-1 do if (G[ig].ColFormats[ii]='#,##0.00') and (Valeur(G[ig].Cells[ii,i])=0.0) then G[ig].Cells[ii,i]:='';
             end;
         end;
    end;
if nbVisible<> 0 then Hauteur:=Trunc((Ecran.ClientHeight-35)/nbVisible) else begin Hauteur:=Trunc(Ecran.ClientHeight-35); {PGIinfo('Aucun détail à afficher',Ecran.Caption); PostMessage(TFVierge(Ecran).Handle,WM_CLOSE,0,0) ; exit; }end;;
for ig:=0 to  MAXG do
    begin
    if GetControlVisible('P'+intToStr(ig)) then SetControlProperty('P'+intToStr(ig),'Height',Hauteur);
    if (Ig>0) and (Ig<MAXG) then SetControlProperty('PBASE'+intToStr(Ig),'Height',(Ecran.ClientHeight-35)-(Ig*Hauteur));
    end;
end;

Procedure TOF_GCENCOURSDETGC.AfficherJustif ;
begin
Reaffiche:=true;
OnLoad;
end;

Procedure TOF_GCENCOURSDETGC.AddSplitter(Ig :Integer) ;
var     PP:THPanel;
begin
   Splits[Ig]:=TSplitter.Create(Ecran);
   Splits[Ig].Name:='Splitter'+intToStr(Ig);
   PP:=THPanel(GetControl('P'+intToStr(Ig))) ;
   Splits[Ig].Parent:=THPanel(GetControl('PBASE'+intToStr(Ig))) ;
   Splits[Ig].AutoSnap:=False;
   Splits[Ig].Beveled:=True;
   Splits[Ig].ResizeStyle:=rsUpdate;
   Splits[Ig].Left:=PP.Left ;
   Splits[Ig].Align:=alTop;
   Splits[Ig].Width:=PP.Width;
   Splits[Ig].Height:=3;
   Splits[Ig].Cursor:=crVSplit;
   Splits[Ig].Color:=clActiveCaption;
end;

Procedure TOF_GCENCOURSDETGC.OnClose ;
var Ig : integer;
begin
inherited ;
For Ig:=0 to MAXG do begin TobA[ig].free; Splits[Ig].free; end;
end;

procedure TOF_GCENCOURSDETGC.GSLigneDClick (Sender: TObject);
var TobP : Tob ;
    CleDoc : R_CleDoc;
    Ig : integer;
begin
if not (sender is thgrid) then exit;
GS:=THGrid(Sender);
Ig:=strToInt(copy(GS.name,2,1));
TobP:=Tob(GS.Objects[0,GS.Row]) ;
if TOBP=nil then exit;
if Ig=0 then
   begin
   CleDoc.NaturePiece:=Uppercase(TOBP.GetValue('GP_NATUREPIECEG'));
   CleDoc.Souche:=Uppercase(TOBP.GetValue('GP_SOUCHE'));
   CleDoc.DatePiece:=TOBP.GetValue('GP_DATEPIECE');
   CleDoc.NumeroPiece:=TOBP.GetValue('GP_NUMERO');
   CleDoc.Indice:=TOBP.GetValue('GP_INDICEG');
   SaisiePiece(CleDoc,taConsult);
   end;

if Ig=1 then
   begin
   AfficherECdetail (TobP);
   end ;


if Ig=2 then
   begin
   if (TOBP.GetValue('E_REFGESCOM') <> '') then
      begin
      CleDoc := DecodeRefCPGescom (TOBP.GetValue('E_REFGESCOM'));
      SaisiePiece(CleDoc,taConsult);
      end;
   end ;
// cd 050701

end;

// cd 050701
Procedure TOF_GCENCOURSDETGC.AfficherEC ;
var TobP : Tob ;
begin
if (TOBA[2] = Nil) or (TOBA[2].Detail.Count = 0) then exit;
TobP:=Tob(G[2].Objects[0,G[2].Row]) ;
if TOBP=nil then exit;
AfficherECdetail (TobP);
end;

Procedure TOF_GCENCOURSDETGC.AfficherECDetail (TobP : TOB) ;
var XX : RMVT ;
begin
   FillChar(XX,Sizeof(XX),#0) ;
   XX.Jal:=TOBP.GetValue('E_JOURNAL');
   XX.DateC:=TOBP.GetValue('E_DATECOMPTABLE');
   XX.Exo:=QuelExoDT(XX.DateC) ;
   XX.Num:=TOBP.GetValue('E_NUMEROPIECE');
   XX.Simul:=TOBP.GetValue('E_QUALIFPIECE');
   XX.Nature:=TOBP.GetValue('E_NATUREPIECE');
   XX.CodeD:=TOBP.GetValue('E_DEVISE') ;
   XX.Etabl:=TOBP.GetValue('E_ETABLISSEMENT') ;
   LanceZoomBordereau(XX)
end;
// cd 050701

procedure TOF_GCENCOURSDETGC.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if sender is THGrid then THGrid(Sender).InvalidateRow(ou) ;
end;

procedure TOF_GCENCOURSDETGC.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if sender is THGrid then THGrid(Sender).InvalidateRow(ou) ;
end;
{==============================================================================================}
{=================================== Impression détail ========================================}
{==============================================================================================}
function TOF_GCENCOURSDETGC.PrepareImpression : integer ;
BEGIN
Result:=0;
END;

procedure TOF_GCENCOURSDETGC.BImprimerClick(Sender: TObject);
var stTiers : string ;
begin
PEtat.Tip:='E'; PEtat.Titre:=TFVierge(Ecran).Caption;
PEtat.Apercu:=True; PEtat.DeuxPages:=False; PEtat.First:=True;
stTiers := 'T_TIERS = "'+TOBTiers.GetValue('T_TIERS')+'"' ;
PEtat.stSQL:=stTiers;
PEtat.Nat:='ENC';
PEtat.Modele:='ENC';

EntreeOptionEdit(PEtat.Tip,PEtat.Nat,PEtat.Modele,PEtat.Apercu,PEtat.DeuxPages,PEtat.First,
                 TPageControl(GetControl('Pages')),PEtat.stSQL,PEtat.Titre, PrepareImpression);

end;

//=================================================================================================

Procedure RisqueTiersGCDetail ( TOBTiers,TobPieces : TOB );
Var StNat,SQL,CodeTiers : String ;
    i     : integer ;
    TOBPP : TOB ;
    Q     : TQuery ;
BEGIN
StNat:='' ;
If TobPieces=Nil then Exit;
if TOBTiers=Nil then Exit ;
CodeTiers:=TOBTiers.GetValue('T_TIERS') ; if CodeTiers='' then Exit ;
for i:=0 to VH_GC.TOBParPiece.Detail.Count-1 do
    BEGIN
    TOBPP:=VH_GC.TOBParPiece.Detail[i] ;
    if ((TOBPP.GetValue('GPP_ENCOURS')='X') and (TOBPP.GetValue('GPP_TYPEECRCPTA')<>'NOR') and
        (TOBPP.GetValue('GPP_VENTEACHAT')='VEN'))
       then StNat:=StNat+' OR GP_NATUREPIECEG="'+TOBPP.GetValue('GPP_NATUREPIECEG')+'"' ;
    END ;
if StNat='' then Exit ;
Delete(StNat,1,4) ; StNat:='('+StNat+')' ;
SQL:='SELECT * FROM PIECE WHERE GP_TIERS="'+CodeTiers+'" AND GP_VIVANTE="X" AND '+StNat ;
Q:=OpenSQL(SQL,True) ;
if Not Q.EOF then TobPieces.loadDetailDB('PIECE','','',Q, False);
Ferme(Q) ;
END ;

Procedure EnCoursRegleNonEchuDetail (Aux : String ; DateButoir : tDateTime; TobNonECh : TOB);
Var St,St1 : String ;
    Q : TQuery ;
BEGIN
if TOBNonEch=Nil then Exit ;
st:='SELECT E_DATEECHEANCE,E_DATECOMPTABLE,E_JOURNAL,E_NUMEROPIECE, E_LIBELLE,E_MODEPAIE, (E_DEBIT-E_CREDIT) as Solde  ';
st:=st+' ,E_QUALIFPIECE,E_NATUREPIECE,E_DEVISE,E_ETABLISSEMENT,E_REFINTERNE ' ;
st:=st+' FROM ECRITURE ' ;
St:=St+' LEFT JOIN MODEPAIE ON E_MODEPAIE=MP_MODEPAIE ' ;
St:=St+' LEFT JOIN GENERAUX ON E_CONTREPARTIEGEN=G_GENERAL ' ;
St:=St+' WHERE E_AUXILIAIRE="'+Aux+'" ' ;
St:=St+' AND (MP_CATEGORIE="LCR") AND (G_NATUREGENE="BQE" OR G_NATUREGENE="CAI") '  ;
St:=st+' AND E_DATEECHEANCE>"'+USDATETIME(DateButoir)+'" ' ;
St:=St+' AND E_QUALIFPIECE="N" AND E_MODEPAIE<>"" AND (E_ECRANOUVEAU="N" OR (E_ECRANOUVEAU="H" AND E_EXERCICE="'+VH^.ExoV8.Code+'"))' ;
St1:=LWhereV8 ; If St1<>'' Then St:=St+' AND '+St1 ;
st:=st+'   Order By E_DATEECHEANCE,E_DATECOMPTABLE,E_NUMEROPIECE';
Q:=OpenSQL(St,TRUE) ;
if Not Q.EOF then TOBNonEch.loadDetailDB('ECRITURE','','',Q, False);
Ferme(Q) ;
END ;


Procedure EnCoursJustifSoldeDetail (Aux : String ; DateButoir : tDateTime; TobSolde : TOB; TypeEcr : string);
Var St,St1 : String ;
    Q : TQuery ;
    SoldeLettre : double ;
    i : integer ;
    TOBE:TOB;
BEGIN
if TOBSolde=Nil then Exit ;
st:='Select E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE ';
st:=st+' ,E_REFINTERNE,  E_LIBELLE, E_ETATLETTRAGE, E_JOURNAL ';
st:=st+' ,E_DATEECHEANCE, E_MODEPAIE, E_DATEPAQUETMAX, E_EXERCICE';
st:=st+' ,E_DEBIT, E_CREDIT, E_COUVERTURE ,E_LETTRAGE';
st:=st+' ,E_QUALIFPIECE,E_NATUREPIECE,E_DEVISE,E_ETABLISSEMENT ' ;
// cd 050701
st:=st+' ,E_REFGESCOM ' ;
// cd 050701
st:=st+' From ECRITURE ' ;
st:=st+' Left Join GENERAUX on G_GENERAL=E_GENERAL';
st:=st+' Where E_AUXILIAIRE="'+Aux+'"';
st:=st+'   And (E_DATECOMPTABLE<="'+USDATETIME(DateButoir)+'" Or E_DATEPAQUETMIN<="'+USDATETIME(DateButoir)+'")';
st:=st+'   And (E_QUALIFPIECE="N" And (E_QUALIFPIECE<>"R")) ';
if TypeEcr='N' then st:=st+'   And E_ETATLETTRAGE<>"TL"'
else if TypeEcr='L' then st:=st+'   And E_ETATLETTRAGE="TL"' ;
st:=st+'   And E_ETATLETTRAGE<>"RI" And E_ECRANOUVEAU<>"CLO" And E_ECRANOUVEAU<>"OAN" and E_QUALIFPIECE<>"C"';
St1:=LWhereV8 ; If St1<>'' Then St:=St+' AND '+St1 ;
st:=st+'   Order By E_ETATLETTRAGE DESC ,E_LETTRAGE,E_DATECOMPTABLE,E_NUMEROPIECE';
Q:=OpenSQL(St,TRUE) ;
if Not Q.EOF then TOBSolde.loadDetailDB('ECRITURE','','',Q, False);
Ferme(Q) ;
if TOBSolde.Detail.Count >0 then
   for i:=0 to TOBSolde.Detail.Count-1 do
       begin
       TOBE:=TOBSolde.Detail[i];
       TOBE.addchampsup('Soldelettre',false) ;
       if TOBE.GetValue('E_DEBIT')=0 then SoldeLettre:= TOBE.GetValue('E_COUVERTURE')-TOBE.GetValue('E_CREDIT')
                                         else SoldeLettre:= TOBE.GetValue('E_DEBIT')-TOBE.GetValue('E_COUVERTURE');
       TOBE.putValue('Soldelettre',SoldeLettre);
       end;
end;

Function EnCoursEchuNonRegle (Aux : String ; DateButoir : tDateTime ): double;
Var St,St1 : String ;
    Q : TQuery ;
BEGIN
result:=0;
st:='Select sum(E_DEBIT-E_COUVERTURE)';
st:=st+' From ECRITURE ' ;
st:=st+' Left Join GENERAUX on G_GENERAL=E_GENERAL';
st:=st+' Where E_AUXILIAIRE="'+Aux+'" and E_DEBIT>0 ';
st:=st+'   And (E_DATECOMPTABLE<="'+USDATETIME(DateButoir)+'" Or E_DATEPAQUETMIN<="'+USDATETIME(DateButoir)+'")';
St:=st+'   AND E_DATEECHEANCE<"'+USDATETIME(DateButoir)+'" ' ;
st:=st+'   And E_QUALIFPIECE="N"  And E_ETATLETTRAGE<>"TL"';
st:=st+'   And E_ETATLETTRAGE<>"RI" And E_ECRANOUVEAU<>"CLO" And E_ECRANOUVEAU<>"OAN" and E_QUALIFPIECE<>"C"';
St1:=LWhereV8 ; If St1<>'' Then St:=St+' AND '+St1 ;
Q:=OpenSQL(St,TRUE) ;
if Not Q.EOF then Result:=Q.Fields[0].AsFloat ;
Ferme(Q) ;
end;


procedure DessineCell(GS:THGrid; ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
  Arect: Trect ;
begin
If Arow < GS.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=GS.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := GS.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = GS.row) then
         BEGIN
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=ARect.Right-2 ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         END ;
    end;
end;


procedure AGLGCENCOURS_AfficherDetail( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  TOF_GCENCOURS(totof).AfficherDetail (string(Parms[1]));
end;

procedure AGLGCENCOURS_AfficherJustif( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  TOF_GCENCOURSDETGC(totof).AfficherJustif;
end;

// cd 050701
procedure AGLGCENCOURS_AfficherEC( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  TOF_GCENCOURSDETGC(totof).AfficherEC;
end;
// cd 050701

Initialization
  registerclasses ( [ TOF_GCENCOURS,TOF_GCENCOURSDETGC ] ) ;
  RegisterAglProc( 'GCENCOURS_AfficherDetail', TRUE , 1, AGLGCENCOURS_AfficherDetail);
  RegisterAglProc( 'GCENCOURS_AfficherJustif', TRUE , 1, AGLGCENCOURS_AfficherJustif);
// cd 050701
  RegisterAglProc( 'GCENCOURS_AfficherEC', TRUE , 1, AGLGCENCOURS_AfficherEC);
// cd 050701

end.


