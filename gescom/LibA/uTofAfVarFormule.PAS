{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVFORMULE ()
Mots clfs ... : TOF;uTofAfVarFormule
*****************************************************************}
Unit uTofAfVarFormule ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
      Fe_Main,
{$Else}
     MainEagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,Utob,HTB97,windows,messages,extctrls,dicoaf,Graphics;

Type

TOF_AFFORMULEVAR = Class (TOF)
 private
    FormuleOk : Boolean ;
    Action   : TActionFiche ;
    Clause_where : String ;
    TobFormule : Tob ;
    TobFormuledetail : Tob ;
    TobInsert : Tob  ;
    QFormule : Tquery ;
    FrappeClavier,OldFrappeClavier : String ;

    AVF_QTETARIFFOR : TRadioButton ;
    AVF_QTETARIFFINALE : TRadioButton ;
    AVF_QTETARIFPRINC : TRadioButton ;

    LBFRAPPECLAVIER : THlabel ;
    BtnAjouteUnite : TToolbarButton97 ;
    BtnRetrancheUnite: TToolbarButton97 ;
    BTNPAROUV : TToolbarButton97 ;
    BTNPARFER: TToolbarButton97 ;
    BTNPOINTVIRGULE :TToolbarButton97 ;
    Btn1: TToolbarButton97 ;
    Btn2: TToolbarButton97 ;
    Btn3: TToolbarButton97 ;
    Btn4: TToolbarButton97 ;
    Btn5: TToolbarButton97 ;
    Btn6: TToolbarButton97 ;
    Btn7: TToolbarButton97 ;
    Btn8: TToolbarButton97 ;
    Btn9: TToolbarButton97 ;
    Btn0: TToolbarButton97 ;
    BEFF :TToolbarButton97 ;

    BEFFMESURE :TToolbarButton97 ;

    BTNDECIMALSEPARATOR: TToolbarButton97 ;
    BTNPLUS: TToolbarButton97 ;
    BTNMOINS: TToolbarButton97 ;
    BTNFOIS: TToolbarButton97 ;
    BTNDIV: TToolbarButton97 ;
    bValider : TToolbarButton97 ;
    bFerme : TToolbarButton97 ;

    Binsert : TToolbarButton97 ;
    Bdelete : TToolbarButton97 ;
    Bdefaire  :TToolbarButton97 ;
    BVerification :TToolbarButton97 ;
    Bduplication:TToolbarButton97 ;
    MonTimer : TTimer;
    GridUnite : THGrid ;
    Griddetail : THGrid ;
    MemoFormule : TMemo ;

    procedure RempliListeUnite ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    procedure BtnAjouteUniteClick(sender : Tobject) ;
    procedure BtnRetrancheUniteClick(sender : Tobject) ;



    procedure BTNPAROUVClick(sender : Tobject) ;
    procedure BTNPARFERClick(sender : Tobject) ;
    procedure Btn1Click(sender : Tobject) ;
    procedure Btn2Click(sender : Tobject) ;
    procedure Btn3Click(sender : Tobject) ;
    procedure Btn4Click(sender : Tobject) ;
    procedure Btn5Click(sender : Tobject) ;
    procedure Btn6Click(sender : Tobject) ;
    procedure Btn7Click(sender : Tobject) ;
    procedure Btn8Click(sender : Tobject) ;
    procedure Btn9Click(sender : Tobject) ;
    procedure Btn0Click(sender : Tobject) ;
    procedure BEFFClick(sender : Tobject) ;
    procedure BEFFMESUREClick(sender : Tobject) ;
    procedure BTNDECIMALSEPARATORClick(sender : Tobject) ;
    procedure BTNPLUSClick(sender : Tobject) ;
    procedure BTNMOINSClick(sender : Tobject) ;
    procedure BTNFOISClick(sender : Tobject) ;
    procedure BTNDIVClick(sender : Tobject) ;
    procedure bValiderClick(sender : Tobject) ;
    procedure bFermeClick(sender : Tobject) ;
    procedure BinsertClick(sender : Tobject) ;
    procedure BDeleteClick(sender : Tobject) ;
    procedure BdefaireClick(sender : Tobject) ;
    procedure BVerificationClick(sender : Tobject) ;
    procedure BDuplicationClick(sender : Tobject) ;

    procedure RadioBouttonClick(sender : Tobject) ;

    procedure GridUniteKeyPress(Sender: TObject; var Key: Char);
    procedure MemoFormuleKeyPress(Sender: TObject; var Key: Char);
    procedure GridUniteDblClick(Sender: TObject);
    procedure VideFrappeClavier(Sender : Tobject) ;
    procedure EcritFormuleLib ;
    procedure VideChamps ;
    procedure ChargeChamps ;
    procedure MontreBouton;
    procedure FabriqueLesTobs;
    procedure FabriqueLaTobDetail;
    function VerifUnePaireCrochets(StUnite :String) : boolean ;
    function VerifContenuCrochets(St : string) : integer ;
    function VerifParenthese(St : String)  : integer;
    function VerifOperateur(St : String)  : integer;
    function TrouveUneLigne(formule,unite : string ; Rang : integer) : integer ;
    function RempliUneLigne(StUnite,StLibelle  :string) : boolean;
    function Remplilagrille(StUnite  :string ) : boolean;
    Function  FabriqueGrille : boolean  ;
    procedure Verification (FormuleTarifaire : boolean) ;
    function VerifieFormule(FormuleTarifaire : boolean) : integer ;
    function  VerifAuMoinsUnOperateur(St : String)  : integer;
    function verificationGrille : integer ;
    procedure MajDetailFormule ;
    function  OkPourModifSupp : boolean ;
    procedure Enregistre(EnSilence : boolean) ;
    function  VerifVide(St : String)  : integer;
    Function OkFormuleTarifaire : boolean ;
    Function OkPremiereQte : boolean ;
    Function TrouvePremiereUnite : String ;
   end ;

Const
TexteMessage: array[1..2] of string 	= (
{1}  'Confirmez-vous la Suppression de cette formule ? ',
{2}  'Vérification effectuée avec succès !'
)  ;

TextedesErreurs: array[1..13] of string 	= (
{1}  'il manque un ].',
{2}  'il manque un [.' ,
{3}  'La formule n''est pas valide!' ,
{4}  'Cette formule ne peut etre modifiée, elle est utilisée dans une affaire .' ,
{5}  'Formule Tarifaire : ' ,
{6}  'Seules les valeurs O,N sont acceptées dans la grille',
{7}  'Une Unité n''est pas reconnu dans la formule.',
{8}  'Il manque une ( .',
{9}  'Il manque une ) .',
{10} 'Incohérence avec le parenthésage ! ',
{11} 'Une Formule ne peut se terminer par un opérateur! ',
{12} 'Une Formule ne peut être vide',
{13} 'Une Formule doit contenir au moins un opérateur!')  ;





procedure AFLanceFiche_FormuleVariable (cle,Action : string ) ;

Implementation
procedure TOF_AFFORMULEVAR.BDuplicationClick(sender : Tobject) ;
begin
  Action:=taCreat ;
  setcontroltext('AVF_FORMULEVAR',getcontroltext('AVF_FORMULEVAR')+'_') ;
  MontreBouton;
end ;



function TOF_AFFORMULEVAR.VerifUnePaireCrochets(StUnite  :String) : boolean ;
var i : integer ;
begin
  i:=0 ;
  while ((i<GridUnite.rowCount) and ((uppercase(GridUnite.Cells[0,i])<>StUnite))) do   inc(i) ;
  result:=(i<GridUnite.rowCount) ;
end ;




function TOF_AFFORMULEVAR.VerifContenuCrochets(St : string) : integer ;
var StUnite : String ;
PosOuv,PosFer : integer ;
Begin
  result:=0 ;
  While (pos('[',st)>0) and (result=0)   do
    begin
    PosOuv:=pos('[',st) ;
    PosFer:=pos(']',st) ;
    if PosFer>0 then
      begin
      StUnite:=copy(St,PosOuv+1,PosFer-posOuv-1)  ;
      St[PosFer]:='|';
      if Not (VerifUnePaireCrochets(uppercase(StUnite))) then result:=-7 ;
      end
      Else result:=-1 ;//il manque un crochet fermant
    St[PosOuv]:='|';
    end ;
  if result=0 then if pos(']',st)>0 then result:=-2 ;//il manque un crochet ouvrant
end ;




function TOF_AFFORMULEVAR.VerifParenthese(St : String)  : integer;
Var PosParOuv,PosParFer : integer ;
begin
  result:=0 ;
  While (pos('(',st)>0) and (result=0) do
    begin
    PosParOuv:=pos('(',st) ;
    PosParFer:=pos(')',st) ;
    if (PosParFer>0) then St[PosParFer]:='.' else result:=-9 ;
    if (result=0) and (PosParFer<PosParOuv) then result:=-10 ;
    St[PosParOuv]:='|' ;
    end ;
  if (result=0) then
  if (pos(')',st)>0) then result:=-8 ;
end ;


function TOF_AFFORMULEVAR.VerifVide(St : String)  : integer;
begin
  result:=0 ;
  if st='' then result:=-12 ;
end ;


function TOF_AFFORMULEVAR.VerifOperateur(St : String)  : integer;
begin
  result:=0 ;
  if st<>'' then
    if St[length(st)] in ['*','+','-','/'] then result:=-11 ;
end ;

function TOF_AFFORMULEVAR.VerifAuMoinsUnOperateur(St : String)  : integer;
var i : integer ;
begin
  result:=-13 ;
  for i:=1 to length(st) do
     if St[i] in ['*','+','-','/'] then result:=0 ;
end ;


function TOF_AFFORMULEVAR.verifieFormule(FormuleTarifaire : boolean) : integer ;
Var St : String ;
UneErreur: integer ;
begin
  if FormuleTarifaire then
    st:=getcontrolText('AVF_FORMULETARIF')
  else
    st:=MemoFormule.Lines.CommaText ;
  UneErreur:=VerifContenuCrochets(St) ;
  if UneErreur=0 then UneErreur:=VerifParenthese(St) ;
  if UneErreur=0 then UneErreur:=VerifOperateur(St) ;
  if UneErreur=0 then UneErreur:=VerifVide(St) ;
  if UneErreur=0 then UneErreur:=VerifAuMoinsUnOperateur(St) ;
  result:=UneErreur ;
end ;




procedure TOF_AFFORMULEVAR.verification(FormuleTarifaire : boolean) ;
Var lemessage : String ;
UneErreur: integer ;
begin
UneErreur:=verifieFormule(FormuleTarifaire) ;
  if UneErreur<0 then
    begin
      if FormuleTarifaire then
        lemessage:=TextedesErreurs[5]+TextedesErreurs[-UneErreur]
        else
        lemessage:=TextedesErreurs[-UneErreur] ;
      pgiBoxAF(Lemessage,Ecran.Caption)
    end
    Else
    begin
    //Success
    if not FormuleTarifaire then // dans ce cas je ne veux pas etre averti que c'est ok.
      begin
        PgiinfoAF(TexteMessage[2],Ecran.Caption)
      end ;
    end
end ;

procedure TOF_AFFORMULEVAR.BVerificationClick(sender : Tobject) ;
begin
enregistre(False) ; // j'enregistre sans sortir avec message d'alerte
TobFormuleDetail.free ;  // je detruis la tobFille
FabriqueLaTobDetail ;  // je relie les valeurs
FabriqueGrille ;      // je les affiche
end ;


procedure TOF_AFFORMULEVAR.VideChamps ;
begin
  setcontroltext('AVF_FORMULEVAR','') ;
  setcontroltext('AVF_FORMVARLIB','') ;
  setcontroltext('AVF_FORMVARDEF','') ;
  setcontroltext('AVF_FORMVARACTIVE','') ;
  setcontroltext('AVF_MESURE','') ;
  setcontroltext('AVF_QTETARIFPRINC','') ;
  setcontroltext('AVF_QTETARIFFINALE','') ;
  setcontroltext('AVF_QTETARIFFOR','') ;
  setcontroltext('AVF_FORMULETARIF','') ;
end ;

procedure TOF_AFFORMULEVAR.BinsertClick(sender : Tobject) ;
begin
  VideChamps ;
  Action:=taCreat ;
  MontreBouton;
  TobFormuleDetail.free ;  // je detruis la tobFille
  FabriqueLaTobDetail ;  // je relie les valeurs
  FabriqueGrille ;      // je les affiche
end ;

function TOF_AFFORMULEVAR.OkPourModifSupp : boolean ;
begin
// verifer si Article, ligne ou AFQTEVAR
result:=True ;
end ;

procedure TOF_AFFORMULEVAR.BDeleteClick(sender : Tobject) ;
begin
  if OkPourModifSupp then
    begin
    if pgiAskCancelAf(TexteMessage[1],ecran.Caption) <> mrYes then Exit;
    ExecuteSQL('DELETE FROM AFORMULEVAR WHERE AVF_FORMULEVAR="'+TobFormule.Detail[0].GetValue('AVF_FORMULEVAR')+'"') ;
    ExecuteSQL('DELETE FROM AFORMULEVARDET WHERE AVD_FORMULEVAR="'+TobFormule.Detail[0].GetValue('AVF_FORMULEVAR')+'"') ;
    BinsertClick(Nil) ;
    end ;
end ;

procedure TOF_AFFORMULEVAR.BdefaireClick(sender : Tobject) ;
begin
  ChargeChamps  ;
end ;

function TOF_AFFORMULEVAR.TrouveUneLigne(formule,unite : string ; Rang : integer) : integer ;
var i : integer ;
begin
// lors de l'initialisation je vais chercher dans les fichiers les valeur des aformulevardets
// si rien dans TobFormuleDetail alors renvoie -1
result:=-1 ;
if (Rang<tobFormuledetail.Detail.Count) then // je vais essayer de trouver la meme unite au meme rang
  begin
    if  (tobFormuledetail.Detail[Rang].GetValue('AVD_FORMULEVAR') = formule) and
        (tobFormuledetail.Detail[Rang].GetValue('AVD_MESURE') = unite) then result:=Rang ;
  end ;
i:=0 ;
// au pire si au meme rang je ne trouve pas de valeur je vais voir autre par.
while ((result<0) and (i<tobFormuledetail.Detail.Count))  do
  begin
  if  (tobFormuledetail.Detail[i].GetValue('AVD_FORMULEVAR') = formule) and
      (tobFormuledetail.Detail[i].GetValue('AVD_MESURE') = unite) then result:=i ;
  inc(i) ;
  end ;
end ;

function PositionneBooleenGridToBase(st : String)  :string ;
begin
st:=uppercase(st) ;
result:='-' ;
if (st='O') or (st='OUI') or (st='X') then result:='X' ;
end ;

function PositionneBooleenBaseToGrid(st : String)  :string ;
begin
st:=uppercase(st) ;
result:='Non' ;
if (st='X')  then result:='Oui' ;
end ;




function TOF_AFFORMULEVAR.RempliUneLigne(StUnite,StLibelle  :string ) : boolean;
var i : integer ;
begin
(* Je considere que la formule a la priorité sur aformlevardet c'est pourquoi je donne priorité a
l'occurence de l'unite dans la formule plutot qu'a son rang dans aformulevardet
Fonctionnellement c'est plus ergonomique => si user deplace une unité dans formule ses saisies restent valides.
*)
GridDetail.Cells[0,Griddetail.RowCount-1]:=StUnite  ;
GridDetail.Cells[1,Griddetail.RowCount-1]:=StLibelle  ;
i:=TrouveUneLigne(GetControlText('AVF_FORMULEVAR'),StUnite,Griddetail.RowCount-2) ; // indice commence a Zero et pas a 2 comme la grille
if i>=0 then
  begin    // au chargement je mets ma Tob dans la grille ouille ! ca fait mal
  GridDetail.Cells[2,Griddetail.RowCount-1]:=tobFormuledetail.Detail[i].GetValue('AVD_LIBMESURECOURT')  ;
  GridDetail.Cells[3,Griddetail.RowCount-1]:=tobFormuledetail.Detail[i].GetValue('AVD_VALDEF')  ;
  GridDetail.Cells[4,Griddetail.RowCount-1]:=PositionneBooleenBaseToGrid(tobFormuledetail.Detail[i].GetValue('AVD_LPVIS'))  ;
  GridDetail.Cells[5,Griddetail.RowCount-1]:=PositionneBooleenBaseToGrid(tobFormuledetail.Detail[i].GetValue('AVD_LPENABLE'))  ;
  GridDetail.Cells[6,Griddetail.RowCount-1]:=PositionneBooleenBaseToGrid(tobFormuledetail.Detail[i].GetValue('AVD_ACTVIS'))  ;
  GridDetail.Cells[7,Griddetail.RowCount-1]:=PositionneBooleenBaseToGrid(tobFormuledetail.Detail[i].GetValue('AVD_ACTENABLE'))  ;
  GridDetail.Cells[8,Griddetail.RowCount-1]:=PositionneBooleenBaseToGrid(tobFormuledetail.Detail[i].GetValue('AVD_PLAVIS'))  ;
  GridDetail.Cells[9,Griddetail.RowCount-1]:=PositionneBooleenBaseToGrid(tobFormuledetail.Detail[i].GetValue('AVD_PLAENABLE'))  ;
  end else
  begin
  GridDetail.Cells[2,Griddetail.RowCount-1]:='' ;
  GridDetail.Cells[3,Griddetail.RowCount-1]:='0'  ;
  GridDetail.Cells[4,Griddetail.RowCount-1]:=PositionneBooleenBaseToGrid('X') ;
  GridDetail.Cells[5,Griddetail.RowCount-1]:=PositionneBooleenBaseToGrid('X') ;
  GridDetail.Cells[6,Griddetail.RowCount-1]:=PositionneBooleenBaseToGrid('X') ;
  GridDetail.Cells[7,Griddetail.RowCount-1]:=PositionneBooleenBaseToGrid('X') ;
  GridDetail.Cells[8,Griddetail.RowCount-1]:=PositionneBooleenBaseToGrid('X') ;
  GridDetail.Cells[9,Griddetail.RowCount-1]:=PositionneBooleenBaseToGrid('X') ;
  end ;
Griddetail.RowCount:=Griddetail.RowCount+1;
end ;


function TOF_AFFORMULEVAR.Remplilagrille(StUnite  :string) : boolean;
var i : integer ;
begin
  i:=0 ;
  while ((i<GridUnite.rowCount) and (uppercase(GridUnite.Cells[0,i])<>StUnite)) do inc(i) ;
    if i<GridUnite.rowCount then
      result:=RempliUneLigne(GridUnite.Cells[0,i],GridUnite.Cells[1,i]) else
      result:=false ;
end ;

Function  TOF_AFFORMULEVAR.FabriqueGrille: boolean  ;
var st,StUnite : String ;
i,PosOuv,PosFer : integer ;
begin
  Griddetail.RowCount:=2 ;
  for i:=0 to 2 do Griddetail.cells[i,1]:='' ; //initialisation sinon ancienne valeur pour insert
  for i:=3 to Griddetail.ColCount-1 do Griddetail.cells[i,1]:='Oui' ; //initialisation sinon ancienne valeur pour insert

  st:=MemoFormule.Lines.CommaText ;
  result:=True ;
  While (pos('[',st)>0) do
    begin
    PosOuv:=pos('[',st) ;
    PosFer:=pos(']',st) ;
    if PosFer>0 then
      begin
      StUnite:=copy(St,PosOuv+1,PosFer-posOuv-1)  ;
      St[PosFer]:='|';
      result:=Remplilagrille(uppercase(StUnite)) ;
      end ;
    St[PosOuv]:='|';
    end ;
    // une ligne est rajoutée en trop lors du dernier passage mais si on ne passe jamais (formule vide) il ne faut pas l'enlever
    if Griddetail.RowCount>2 then Griddetail.RowCount:=Griddetail.RowCount-1;
end ;

procedure TOF_AFFORMULEVAR.EcritFormuleLib ;
var i : integer ;
st : String ;
begin
  st:= MemoFormule.text ;
  i:=0 ;
  while (i<GridUnite.rowCount) do
  begin
  St:=StringReplace(St,'['+GridUnite.Cells[0,i]+']',' '+GridUnite.Cells[1,i]+' ', [rfReplaceAll,rfIgnoreCase]) ;
  inc(i) ;
  end ;
  setControlText('AFE_FORMULELIB',st) ;
end ;

procedure TOF_AFFORMULEVAR.BtnAjouteUniteClick(sender : Tobject) ;
begin
  MemoFormule.Seltext:='['+GridUnite.Cells[0,GridUnite.Row]+']' ;
  EcritFormuleLib ;
end ;

procedure TOF_AFFORMULEVAR.GridUniteDblClick(Sender: TObject);
begin
  MemoFormule.Seltext:='['+GridUnite.Cells[0,GridUnite.Row]+']' ;
  EcritFormuleLib ;
end ;

procedure TOF_AFFORMULEVAR.BtnRetrancheUniteClick(sender : Tobject) ;
Var St : String ;
begin
  St:=MemoFormule.text ;
  St:=StringReplace(St,'['+GridUnite.Cells[0,GridUnite.Row]+']','', [rfReplaceAll,rfIgnoreCase]) ;
  MemoFormule.text :=St ;
  EcritFormuleLib ;
end ;


procedure TOF_AFFORMULEVAR.BTNPAROUVClick(sender : Tobject) ;
begin
  MemoFormule.seltext:='(' ;
end ;

procedure TOF_AFFORMULEVAR.BTNPARFERClick(sender : Tobject) ;
begin
  MemoFormule.seltext:=')' ;
end ;

procedure TOF_AFFORMULEVAR.Btn1Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='1' ;
end ;

procedure TOF_AFFORMULEVAR.Btn2Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='2' ;
end ;

procedure TOF_AFFORMULEVAR.Btn3Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='3' ;
end ;

procedure TOF_AFFORMULEVAR.Btn4Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='4' ;
end ;

procedure TOF_AFFORMULEVAR.Btn5Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='5' ;
end ;

procedure TOF_AFFORMULEVAR.Btn6Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='6' ;
  end ;

procedure TOF_AFFORMULEVAR.Btn7Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='7' ;
end ;

procedure TOF_AFFORMULEVAR.Btn8Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='8' ;
end ;

procedure TOF_AFFORMULEVAR.Btn9Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='9' ;
end ;

procedure TOF_AFFORMULEVAR.Btn0Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='0' ;
end ;



procedure TOF_AFFORMULEVAR.BTNDECIMALSEPARATORClick(sender : Tobject) ;
begin
  MemoFormule.Seltext:=DecimalSeparator ;
end ;

procedure TOF_AFFORMULEVAR.BTNPLUSClick(sender : Tobject) ;
begin
  MemoFormule.Seltext:='+';
end ;

procedure TOF_AFFORMULEVAR.BTNMOINSClick(sender : Tobject) ;
begin
  MemoFormule.Seltext:='-';
end ;

procedure TOF_AFFORMULEVAR.BTNFOISClick(sender : Tobject) ;
begin
  MemoFormule.Seltext:='*';
end ;

procedure TOF_AFFORMULEVAR.BTNDIVClick(sender : Tobject) ;
begin
  MemoFormule.Seltext:='/';
end ;

procedure TOF_AFFORMULEVAR.BEFFClick(sender : Tobject) ;
begin
  MemoFormule.text:='';
  setControlText('AFE_FORMULELIB','') ;
end ;

procedure TOF_AFFORMULEVAR.BEFFMESUREClick(sender : Tobject) ;
begin
  SetControlText('AVF_FORMULETARIF','');
end ;


procedure TOF_AFFORMULEVAR.MemoFormuleKeyPress(Sender: TObject; var Key: Char);
begin
  if not (Key in [#22,#3,'(','.',',',')',' ',';',#8,'0'..'9','*','+','-','/']) then key:=#0 ;
end ;


procedure TOF_AFFORMULEVAR.GridUniteKeyPress(Sender: TObject; var Key: Char);
Var i : integer ;
begin
  i:=0 ;
  FrappeClavier:=Uppercase(FrappeClavier+Key) ;
  While (i<GridUnite.RowCount) and (copy(uppercase(GridUnite.Cells[0,i]),1,length(FrappeClavier))<>FrappeClavier) do i:=i+1 ;
  if i>=GridUnite.RowCount then i:=0 ;
  GridUnite.GotoRow(i) ;
  LBFRAPPECLAVIER.caption:=FrappeClavier ;
end ;


procedure TOF_AFFORMULEVAR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFFORMULEVAR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFFORMULEVAR.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFFORMULEVAR.MontreBouton;
begin
  Bdelete.Enabled:=Action=taModif ;
  Binsert.Enabled:=Action=taModif ;
  Bduplication.Enabled:=Action=taModif ;
  SetControlEnabled('AVF_FORMULEVAR',Action=taCreat) ;
end ;

procedure TOF_AFFORMULEVAR.FabriqueLaTobDetail;
Var stDetail : String ;
QFormuleDetail : Tquery ;
begin
   TobFormuledetail:=TOB.Create('Mes details de Formule',nil,-1);
    try
    stDetail:='SELECT AVD_FORMULEVAR, AVD_RANG, AVD_MESURE,AVD_LIBMESURE, AVD_LIBMESURECOURT,' ;
    stDetail:=stDetail+' AVD_VALDEF,AVD_LPVIS, AVD_LPENABLE, AVD_ACTVIS, AVD_ACTENABLE,AVD_PLAVIS,' ;
    stDetail:=stDetail+' AVD_PLAENABLE FROM AFORMULEVARDET inner join  AFORMULEVAR on AVF_FORMULEVAR=AVD_FORMULEVAR ' ;
    stDetail:=stDetail+Clause_where ;
    QFormuleDetail:=Nil ;
    QFormuleDetail := OpenSQL(stDetail, TRUE);
    TobFormuledetail.LoadDetailDB('AFORMULEVARDET','','',QFormuleDetail,false) ;
    finally
    Ferme(QFormuleDetail) ;
    end ; // Try

end ;
procedure TOF_AFFORMULEVAR.FabriqueLesTobs;
Var StQformule : String ;
begin
TobFormule:=TOB.Create('Ma Formule',nil,-1);
  if Action=taCreat then
    begin
    TobInsert:=TOB.Create('AFORMULEVAR',TobFormule,-1);
    end Else
    begin
    try
      StQformule:='SELECT AVF_FORMULEVAR,AVF_FORMULETARIF,AVF_QTETARIFFOR, AVF_QTETARIFFINALE,' ;
      StQformule:=StQformule+'AVF_QTETARIFPRINC, AVF_MESURE, AVF_FORMVARACTIVE,' ;
      StQformule:=StQformule+'AVF_FORMVARDEF, AVF_FORMVARLIB ' ;
      StQformule:=StQformule+'FROM AFORMULEVAR '+Clause_where ;
      QFormule:=Nil ;
      QFormule := OpenSQL(StQformule, TRUE);
      TobFormule.LoadDetailDB('AFORMULEVAR','','',QFormule,false) ;
      FabriqueLaTobDetail ;
      ChargeChamps  ;
    finally
    Ferme(QFormule) ;
    end ; // Try
   end ; // if TaCrea

end ;

procedure TOF_AFFORMULEVAR.ChargeChamps ;
var st : string ;
begin
if TobFormule.Detail.count>0 then
  begin
  setcontroltext('AVF_FORMULEVAR',TobFormule.Detail[0].GetValue('AVF_FORMULEVAR')) ;
  setcontroltext('AVF_FORMVARLIB',TobFormule.Detail[0].GetValue('AVF_FORMVARLIB')) ;
  setcontroltext('AVF_FORMVARDEF',TobFormule.Detail[0].GetValue('AVF_FORMVARDEF')) ;
  SetControlChecked('AVF_FORMVARACTIVE',TobFormule.Detail[0].GetValue('AVF_FORMVARACTIVE')='X')  ;
  TobFormule.Detail[0].GetValue('AVF_MESURE') ;
  setcontroltext('AVF_MESURE',TobFormule.Detail[0].GetValue('AVF_MESURE')) ;
  SetControlChecked('AVF_QTETARIFPRINC',TobFormule.Detail[0].GetValue('AVF_QTETARIFPRINC')='X')  ;
  SetControlChecked('AVF_QTETARIFFINALE',TobFormule.Detail[0].GetValue('AVF_QTETARIFFINALE')='X')  ;
  SetControlChecked('AVF_QTETARIFFOR',TobFormule.Detail[0].GetValue('AVF_QTETARIFFOR')='X')  ;
  setcontroltext('AVF_FORMULETARIF',TobFormule.Detail[0].GetValue('AVF_FORMULETARIF')) ;
  end ;  // if Count
end ;

procedure TOF_AFFORMULEVAR.OnLoad ;
begin
  Inherited ;
  FormuleOk:=True ;
  MontreBouton ;
  FrappeClavier:='';
  OldFrappeClavier:='' ;
  RempliListeUnite ;
  MonTimer:=TTimer.Create(Ecran) ;
  MonTimer.Interval:=500 ;
  MonTimer.OnTimer:=VideFrappeClavier ;
  FabriqueLesTobs ;
//  if OkFormuleTarifaire then verification(true,true) ;// pour formule tarifaire
//  verification(true,false) ;
  verifieFormule(false)  ;

  FabriqueGrille ;

end ;



function valeur(st : string) : string ;
var code : integer ;
f : double ;
begin
val(st,f,code) ;
result:=floattostr(f) ;
end ;

procedure  TOF_AFFORMULEVAR.MajDetailFormule ;
var st,stValues,Formule,mesure : String ;
i,colonne,ligne : integer ;
libMesure,libMesureCourt : String ;
begin
formule:=TobFormule.Detail[0].GetValue('AVF_FORMULEVAR') ;
st:='delete from  AFORMULEVARDET where AVD_FORMULEVAR = "'+formule+'"' ;
executesql(st) ;
st:='insert into AFORMULEVARDET (AVD_FORMULEVAR, AVD_RANG, AVD_MESURE,AVD_LIBMESURE, AVD_LIBMESURECOURT, AVD_VALDEF,' ;
st:=st+'AVD_LPVIS, AVD_LPENABLE, AVD_ACTVIS, AVD_ACTENABLE,AVD_PLAVIS, AVD_PLAENABLE) values (';
for ligne:=1 to GridDetail.rowcount-1 do
  begin
  mesure:=GridDetail.Cells[0,ligne] ;
  libMesure:=GridDetail.Cells[1,ligne] ;
  libMesureCourt:=GridDetail.Cells[2,ligne] ;
  stValues:='"'+formule+'",'+inttostr(ligne)+',"'+mesure+'","'+libMesure+'","'+libMesureCourt+'",'+valeur(GridDetail.Cells[3,ligne]) ;
  for colonne:=4 to 9 do
    stValues:=stValues+',"'+PositionneBooleenGridToBase(GridDetail.Cells[colonne,ligne])+'"' ;
  executesql(st+stValues+')') ;
  end ;
end ;

Function  TOF_AFFORMULEVAR.verificationGrille  : integer ;

var ligne,Colonne : integer ;
erreur : integer ;
begin
erreur:=0 ;
for ligne:=1 to GridDetail.rowcount-1 do
  for colonne:=4 to 9 do
   if (uppercase(GridDetail.Cells[colonne,ligne])<>'O')  and
      (uppercase(GridDetail.Cells[colonne,ligne])<>'OUI') and
      (uppercase(GridDetail.Cells[colonne,ligne])<>'N')  and
      (uppercase(GridDetail.Cells[colonne,ligne])<>'NON') then erreur :=-6 ;
result:=erreur ;
end ;

procedure TOF_AFFORMULEVAR.Enregistre(EnSilence : boolean) ;
Var Erreur  :integer ;
begin
//ExisteSQL ('select AVF_FORMULEVAR from AFORMULEVAR  where AVF_FORMULEVAR="'+TobFormule.Detail[0].GetValue('AVF_FORMULEVAR')+'"')
  // table Article,Ligne,AfqteVar
  if false  then
    pgiBoxAF(TextedesErreurs[4],Ecran.Caption)
  else
  begin
    if OkFormuleTarifaire then  erreur:=verifieFormule(true)  ;
      if erreur=0 then erreur:=verificationGrille ;
      if erreur=0 then erreur:=verifieFormule(false)  ;
    if erreur<>0 then
      begin
      pgiBoxAF(TextedesErreurs[-erreur],Ecran.Caption) ;
      formuleOK:=false ;
      end
      else formuleOK:=true ;
   if (Not Ensilence) and (erreur=0) then PgiinfoAF(TexteMessage[2],Ecran.Caption) ;
   //   HShowMessage('0;'+Ecran.Caption+';'+TextedesErreurs[3]+';W;O;O;O;','','') ;
   if Erreur=0 then
      begin
      TobFormule.Detail[0].PutValue('AVF_FORMULEVAR',GetControltext('AVF_FORMULEVAR')) ;
      TobFormule.Detail[0].PutValue('AVF_FORMVARLIB',GetControltext('AVF_FORMVARLIB')) ;
      TobFormule.Detail[0].PutValue('AVF_FORMVARDEF',GetControltext('AVF_FORMVARDEF')) ;
      TobFormule.Detail[0].PutValue('AVF_FORMVARACTIVE',GetControltext('AVF_FORMVARACTIVE')) ;
      TobFormule.Detail[0].PutValue('AVF_MESURE',GetControltext('AVF_MESURE')) ;
      TobFormule.Detail[0].PutValue('AVF_QTETARIFPRINC',GetControltext('AVF_QTETARIFPRINC')) ;
      TobFormule.Detail[0].PutValue('AVF_QTETARIFFINALE',GetControltext('AVF_QTETARIFFINALE')) ;
      TobFormule.Detail[0].PutValue('AVF_QTETARIFFOR',GetControltext('AVF_QTETARIFFOR')) ;
      TobFormule.Detail[0].PutValue('AVF_FORMULETARIF',GetControltext('AVF_FORMULETARIF')) ;
      if Action=taModif then TobFormule.UpdateDB(False) Else TobFormule.insertDB(nil,false) ;
      MajDetailFormule ;
      Action:=taModif ;
      MontreBouton;
      end ;
    end ;

end  ;

procedure TOF_AFFORMULEVAR.bValiderClick(sender : Tobject) ;
var V : Variant ;
begin
enregistre(true) ;
end ;

procedure TOF_AFFORMULEVAR.bFermeClick(sender : Tobject) ;
var V : Variant ;
begin
FormuleOk:=True ;
end ;

procedure TOF_AFFORMULEVAR.VideFrappeClavier(Sender :Tobject) ;
begin
  if (OldFrappeClavier=FrappeClavier) then FrappeClavier:='';
  OldFrappeClavier:=FrappeClavier;
  LBFRAPPECLAVIER.caption:=FrappeClavier ;
end ;

procedure TOF_AFFORMULEVAR.RempliListeUnite ;
Var StRequete : String ;
    TobUnite:TOB;
    MaRequete : TQuery ;
    i : integer ;

begin
  StRequete:='SELECT GME_MESURE,GME_LIBELLE FROM MEA' ;
  MaRequete:=Nil ;
  TobUnite:=TOB.Create('',nil,-1);
  i:=0 ;
  try
      MaRequete := OpenSQL(StRequete, TRUE);
      TobUnite.LoadDetailDB('','','',MaRequete,false) ;
      GridUnite.RowCount:=TobUnite.Detail.count ;
      while (i<TobUnite.Detail.count) do
        begin
        GridUnite.Cells[0,i]:=TobUnite.Detail[i].GetValue('GME_MESURE')  ;
        GridUnite.Cells[1,i]:=TobUnite.Detail[i].GetValue('GME_LIBELLE')  ;
        i:=i+1 ;
        end ;
      finally
      Ferme(MaRequete);
      TobUnite.Free;
      end ;
end;

Function TOF_AFFORMULEVAR.OkFormuleTarifaire : boolean ;
begin
result:=GetControltext('AVF_QTETARIFFOR')='X' ;
end ;

Function TOF_AFFORMULEVAR.OkPremiereQte : boolean ;
begin
result:=GetControltext('AVF_QTETARIFPRINC')='X' ;
end ;

Function TOF_AFFORMULEVAR.TrouvePremiereUnite : String ;
var StUnite,St : String ;
PosOuv,PosFer : integer ;
begin
st:=MemoFormule.Lines.CommaText ;
PosOuv:=pos('[',st) ;
PosFer:=pos(']',st) ;
StUnite:=copy(St,PosOuv+1,PosFer-posOuv-1)  ;
if VerifUnePaireCrochets(uppercase(StUnite)) then result:=Stunite else result:='' ;
end ;



procedure TOF_AFFORMULEVAR.RadioBouttonClick(sender : Tobject) ;
Var OkMontre : boolean ;
begin
OkMontre:=OkFormuleTarifaire;
setcontrolEnabled('AVF_FORMULETARIF', OKMontre) ;
setcontrolEnabled('BEFFMESURE',OKMontre) ;
if okmontre and (GetControltext('AVF_FORMULETARIF')='') then setControltext('AVF_FORMULETARIF',MemoFormule.Lines.CommaText) ;
if OkPremiereQte and (GetControltext('AVF_MESURE')='') then  SetControlText('AVF_MESURE',TrouvePremiereUnite) ;

if THedit(getcontrol('AVF_FORMULETARIF')).enabled then
   THedit(getcontrol('AVF_FORMULETARIF')).color:=ClWindow
   else
   begin
   THedit(getcontrol('AVF_FORMULETARIF')).color:=clBtnFace	 ;
   SetcontrolText('AVF_FORMULETARIF','') ;
   end ;
end ;



procedure TOF_AFFORMULEVAR.OnArgument (S : String ) ;
  Var  Critere, Champ, valeur  : String;
  X : integer ;
begin
  Inherited ;
  Action:=taCreat;
  Critere:=(Trim(ReadTokenSt(S)));
  While (Critere <>'') do
  Begin
    if Critere<>'' then
    Begin
      X:=pos('=',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if Champ = 'AVF_FORMULEVAR' then Clause_where :='WHERE AVF_FORMULEVAR="'+Valeur+'"';
        if (Champ ='ACTION') and (valeur = 'CREATION') then Action:=taCreat;
        if (Champ ='ACTION') and (valeur = 'MODIFICATION') then Action:=taModif;
        if (Champ ='ACTION') and (valeur = 'CONSULTATION') then Action:=taConsult;
        END;
     Critere:=(Trim(ReadTokenSt(S)));
  END;

  BtnAjouteUnite:=TToolbarButton97(Getcontrol('BTNAJOUTEUNITE')) ;
  BtnRetrancheUnite:=TToolbarButton97(Getcontrol('BTNRETRANCHEUNITE')) ;
  BTNPAROUV:=TToolbarButton97(Getcontrol('BTNPAROUV')) ;
  BTNPARFER:=TToolbarButton97(Getcontrol('BTNPARFER')) ;
  Btn1:=TToolbarButton97(Getcontrol('BTN1')) ;
  Btn2:=TToolbarButton97(Getcontrol('BTN2')) ;
  Btn3:=TToolbarButton97(Getcontrol('BTN3')) ;
  Btn4:=TToolbarButton97(Getcontrol('BTN4')) ;
  Btn5:=TToolbarButton97(Getcontrol('BTN5')) ;
  Btn6:=TToolbarButton97(Getcontrol('BTN6')) ;
  Btn7:=TToolbarButton97(Getcontrol('BTN7')) ;
  Btn8:=TToolbarButton97(Getcontrol('BTN8')) ;
  Btn9:=TToolbarButton97(Getcontrol('BTN9')) ;
  Btn0:=TToolbarButton97(Getcontrol('BTN0')) ;
  BDelete:=TToolbarButton97(Getcontrol('BDELETE')) ;
  Binsert:=TToolbarButton97(Getcontrol('BINSERT')) ;
  Bdefaire:=TToolbarButton97(Getcontrol('BDEFAIRE')) ;
  Binsert.OnClick:=BinsertClick ;
  BDelete.OnClick:=BDeleteClick ;
  Bdefaire.OnClick:=BdefaireClick ;


  BVerification:=TToolbarButton97(Getcontrol('BVERIFICATION')) ;
  BVerification.OnClick:=BVerificationClick;

  BTNDECIMALSEPARATOR:=TToolbarButton97(Getcontrol('BTNDECIMALSEPARATOR')) ;
  BTNPOINTVIRGULE:=TToolbarButton97(Getcontrol('BTNPOINTVIRGULE')) ;

  BTNPLUS:=TToolbarButton97(Getcontrol('BTNPLUS')) ;
  BTNMOINS:=TToolbarButton97(Getcontrol('BTNMOINS')) ;
  BTNFOIS:=TToolbarButton97(Getcontrol('BTNFOIS')) ;
  BTNDIV:=TToolbarButton97(Getcontrol('BTNDIV')) ;
  BEFF:=TToolbarButton97(Getcontrol('BEFF')) ;
  BEFFMESURE:=TToolbarButton97(Getcontrol('BEFFMESURE')) ;
  bValider:=TToolbarButton97(GetControl('bValider')) ;
  bFerme:=TToolbarButton97(GetControl('bFerme')) ;
  BDuplication:=TToolbarButton97(Getcontrol('BDUPLICATION')) ;
  BDuplication.OnClick:=BDuplicationClick ;
  BtnAjouteUnite.OnClick:=BtnAjouteUniteClick ;
  BtnRetrancheUnite.OnClick:=BtnRetrancheUniteClick ;

  BTNPAROUV.OnClick:=BTNPAROUVClick ;
  BTNPARFER.OnClick:=BTNPARFERClick ;

  BEFF.OnClick:=BEFFClick ;
  BEFFMESURE.OnClick:=BEFFMESUREClick ;

  Btn1.OnClick:=Btn1Click ;
  Btn2.OnClick:=Btn2Click ;
  Btn3.OnClick:=Btn3Click ;
  Btn4.OnClick:=Btn4Click ;
  Btn5.OnClick:=Btn5Click ;
  Btn6.OnClick:=Btn6Click ;
  Btn7.OnClick:=Btn7Click ;
  Btn8.OnClick:=Btn8Click ;
  Btn9.OnClick:=Btn9Click ;
  Btn0.OnClick:=Btn0Click ;
  BTNDECIMALSEPARATOR.OnClick:=BTNDECIMALSEPARATORClick ;
  BTNPLUS.OnClick:=BTNPLUSClick ;
  BTNMOINS.OnClick:=BTNMOINSClick ;
  BTNFOIS.OnClick:=BTNFOISClick ;
  BTNDIV.OnClick:=BTNDIVClick ;
  bValider.OnClick:=bValiderClick ;
  bFerme.OnClick:=bFermeClick ;
  LBFRAPPECLAVIER:=THLabel(GetControl('LBFRAPPECLAVIER')) ;
  LBFRAPPECLAVIER.caption:='' ;
  GridUnite:=THGrid(getcontrol('GRIDUNITE'));
  GridDetail:=THGrid(getcontrol('GRIDDETAIL'));

  GridUnite.OnKeyPress:=GridUniteKeyPress;
  GridUnite.OnDblClick:=GridUniteDblClick ;
  MemoFormule:=Tmemo(getControl('AVF_FORMVARDEF')) ;


 AVF_QTETARIFFOR :=TRadioButton(getcontrol('AVF_QTETARIFFOR')) ;
 AVF_QTETARIFFINALE :=TRadioButton(getcontrol('AVF_QTETARIFFINALE')) ;
 AVF_QTETARIFPRINC := TRadioButton(getcontrol('AVF_QTETARIFPRINC')) ;
 AVF_QTETARIFFOR.OnClick:=RadioBouttonClick ;
 AVF_QTETARIFFINALE.OnClick:=RadioBouttonClick ;
 AVF_QTETARIFPRINC.OnClick:=RadioBouttonClick ;

  MemoFormule.OnKeyPress:=MemoFormuleKeyPress ;
end ;

procedure TOF_AFFORMULEVAR.OnClose ;
begin
  Inherited ;
  if formuleOK  then
    begin
    TobFormule.free ;
    TobFormuleDetail.free
    end else
    lastError:=-1 ;
end ;

procedure TOF_AFFORMULEVAR.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFFORMULEVAR.OnCancel () ;
begin
  Inherited ;
end ;
procedure AFLanceFiche_FormuleVariable(cle,Action : string ) ;
begin
  AglLanceFiche ('AFF','AFVARFORMULE','',cle,Action);
end ;


Initialization
  registerclasses ( [ TOF_AFFORMULEVAR ] ) ;
end.
