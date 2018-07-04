{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVFORMULE ()
Mots clfs ... : TOF;AFREVFORMULE
*****************************************************************}
Unit uTofAfRevFormule ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
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
     UTOF,Utob,HTB97,windows,messages,extctrls,formule, Dicobtp,
     uTofAfValIndicemul, uTofAfRevFormuleEdit;

Type

TOF_AFREVFORMULE = Class (TOF)

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure BtnAjouteIndiceClick(sender : Tobject) ;
    procedure BtnAjouteIndice0Click(sender : Tobject) ;
    procedure BtnRetrancheIndiceClick(sender : Tobject) ;
    procedure BtnRetrancheIndice0Click(sender : Tobject) ;
    procedure BTNPAROUVClick(sender : Tobject) ;
    procedure BTNPARFERClick(sender : Tobject) ;
    procedure BTNARRONDIClick(sender : Tobject) ;
    procedure BTNARRONDIFINClick(sender : Tobject) ;
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
//    procedure BEFFClick(sender : Tobject) ;
    procedure BTNDECIMALSEPARATORClick(sender : Tobject) ;
    procedure BTNPOINTVIRGULEClick(sender : Tobject) ;
    procedure BTNPLUSClick(sender : Tobject) ;
    procedure BTNMOINSClick(sender : Tobject) ;
    procedure BTNFOISClick(sender : Tobject) ;
    procedure BTNDIVClick(sender : Tobject) ;
    procedure bValiderClick(sender : Tobject) ;

    procedure BinsertClick(sender : Tobject) ;
    procedure BDeleteClick(sender : Tobject) ;
    procedure BdefaireClick(sender : Tobject) ;

    procedure BDuplicationClick(sender : Tobject) ;
    procedure BVerificationClick(sender : Tobject) ;

    procedure GridIndiceKeyPress(Sender: TObject; var Key: Char);
    procedure MemoFormuleKeyPress(Sender: TObject; var Key: Char);
    procedure BIndiceClick(sender : Tobject);
    procedure BFormuleEditionClick(sender : Tobject);

 private
    FormuleOk : Boolean ;
    interditFermeFenetre : boolean ;
    Action   : TActionFiche ;
    Clause_where : String ;
    TobFormule : Tob ;
    TobInsert : Tob  ;
    QFormule : Tquery ;
    FrappeClavier,OldFrappeClavier : String ;
    LBFRAPPECLAVIER : THlabel ;
    BtnAjouteIndice : TToolbarButton97 ;
    BtnAjouteIndice0: TToolbarButton97 ;
    BtnRetrancheIndice: TToolbarButton97 ;
    BtnRetrancheIndice0: TToolbarButton97 ;
    BTNPAROUV : TToolbarButton97 ;
    BTNPARFER: TToolbarButton97 ;
    BTNARRONDI: TToolbarButton97 ;
    BTNARRONDIFIN: TToolbarButton97 ;
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
//    BEFF :TToolbarButton97 ;
    BTNDECIMALSEPARATOR: TToolbarButton97 ;
    BTNPLUS: TToolbarButton97 ;
    BTNMOINS: TToolbarButton97 ;
    BTNFOIS: TToolbarButton97 ;
    BTNDIV: TToolbarButton97 ;
    bValider : TToolbarButton97 ;
    Binsert : TToolbarButton97 ;
    Bdelete : TToolbarButton97 ;
    Bdefaire  :TToolbarButton97 ;
    BIndice :TToolbarButton97 ;
    BFormuleEdition :TToolbarButton97 ;
    BDuplication :TToolbarButton97 ;
    BVerification :TToolbarButton97 ;
    MonTimer : TTimer;
    GridIndice : THGrid ;
    MemoFormule : Tmemo ;
    procedure RempliListeIndice ;
    procedure VideFrappeClavier(Sender : Tobject) ;
    function  Majindice : boolean ;
    function  TrouveIndiceDansGrid(Stindice  :string) : boolean;
    function  PositionneLabels(StLabel  :string) : boolean;
    procedure InitialiseLabels ;
    procedure VideChamps ;
    procedure ChargeChamps ;
    procedure MontreBouton;
    procedure FabriqueLesTobs;
    function VerifUnePaireCrochets(Stindice  :String) : boolean ;
    function VerifContenuCrochets(St : string) : integer ;
    function VerifArrondi(St : string) : integer ;
    function VerifParenthese(St : String)  : integer;
    function VerifOperateur(St : String)  : integer;
    function VerifIndiceIndice0(St : string)  :  integer ;
    function VerifpasVide(St : string)  :  integer ;

    function PrepareFormuleAvantInterpretation(St : String)  : String;
    function SetValIndiceFormule(pStIndice : HString) : variant;
    function VerifGformule(St : HString)  : integer;

    procedure Verification (EnSilence : boolean) ;
    function verifChampObligatoire : boolean ;
    procedure gerefocus ;
//    function ExisteCoef : Boolean;
    
  end;

Const
fin_indice_initial='°' ;
TexteMessage: array[1..4] of string 	= (
{1}  'Confirmez la suppression de cette formule ? ',
{2}  'Vérification effectuée avec succès !',
{3}  'La formule n''est pas valide!',
{4}  'Cette formule ne peut etre modifiée, un coefficient est appliqué pour une affaire.'
)  ;
TextedesErreurs: array[1..19] of string 	= (
{1}  'il manque un ].',
{2}  'il manque un [.' ,
{3}  'il manque un ,.' ,
{4}  'il manque un }.' ,
{5}  'il y a une , en trop.' ,
{6}  'il y a une } en trop.',
{7}  'Un indice n''est pas reconnu dans la formule.',
{8}  'Il manque une ( .',
{9}  'Il manque une ) .',
{10} 'Incohérence avec le parenthésage ! ',
{11} 'Une Formule ne peut se terminer par un opérateur! ',
{12} 'Il manque un indice0 dans la formule. ',
{13} 'Un indice0 existe dans la formule sans l''indice.',
{14} 'Le mode de lecture des indices est obligatoire.',
{15} 'Le type de la formule est obligatoire.',
{16} 'Le nombre de mois  est obligatoire pour ce mode de lecture d''indice.' ,
{17} 'La formule est obligatoire.',
{18} 'Le code de la formule est obligatoire.',
{19} 'La formule n''est pas égale à 1.'
)  ;



procedure AFLanceFiche_Formule (cle,Action : string ) ;

Implementation

procedure TOF_AFREVFORMULE.BDuplicationClick(sender : Tobject) ;
begin
  Action:=taCreat ;
  setcontroltext('AFE_FORCODE',getcontroltext('AFE_FORCODE')+'_') ;
  MontreBouton;                        
end;

function TOF_AFREVFORMULE.VerifUnePaireCrochets(Stindice  :String) : boolean ;
var i : integer ;
begin
  i:=0 ;
  while ((i<GridIndice.rowCount) and ((uppercase(GridIndice.Cells[0,i])<>Stindice)
                                 and(uppercase(GridIndice.Cells[0,i]+fin_indice_initial)<>Stindice))) do
     inc(i) ;
  result:=(i<GridIndice.rowCount) ;
end;

function TOF_AFREVFORMULE.VerifContenuCrochets(St : string) : integer ;
var Stindice : String ;
PosOuv,PosFer : integer ;
Begin
  result:=0 ;
  While (pos('[',st)>0) and (result=0)   do
    begin
    PosOuv:=pos('[',st) ;
    PosFer:=pos(']',st) ;
    if PosFer>0 then
      begin
      Stindice:=copy(St,PosOuv+1,PosFer-posOuv-1)  ;
      St[PosFer]:='|';
      if Not (VerifUnePaireCrochets(uppercase(Stindice))) then result:=-7 ;
      end
      Else result:=-1 ;//il manque un crochet fermant
    St[PosOuv]:='|';
    end ;
  if result=0 then if pos(']',st)>0 then result:=-2 ;//il manque un crochet ouvrant
end;

function TOF_AFREVFORMULE.VerifIndiceIndice0(St : string)  :  integer ;
var i,PosIndice,PosIndice_o : integer ;
begin
  result:=0;
  i:=0 ;
  while ((i<GridIndice.rowCount)) do begin
  PosIndice:=Pos('['+uppercase(GridIndice.Cells[0,i])+']',St) ;
  PosIndice_o:=Pos('['+uppercase(GridIndice.Cells[0,i]+fin_indice_initial+']'),St) ;
   if (PosIndice>0) and
      (PosIndice_o=0) then result:=-12 ;
   if (PosIndice_o>0) and
      (PosIndice=0) then result:=-13 ;
     inc(i) ;
  end ;
end ;

function TOF_AFREVFORMULE.VerifArrondi(St : string) : integer ;
var PosArr,PosVirgule,PosArrFin : integer ;
Begin
  result:=0  ;
  While (pos('ARR{',st)>0) do
    begin
    PosArr:=pos('ARR{',st) ;
    PosVirgule:=pos(';',st) ;
    if PosVirgule>0 then  St[PosVirgule]:='|' else result:=-3 ;//manque une ,
    PosArrFin:=pos('}',st) ;
    if PosArrFin>0 then  St[PosArrFin]:='|'  else  result:=-4 ; //manque une }
    st[PosArr]:='|' ;
    end ;
  if result=0 then if pos(';',st)>0 then result:=-5 ;  // une , en trop
  if result=0 then if pos('}',st)>0 then result:=-6 ;  //une } en trop
end ;



function TOF_AFREVFORMULE.VerifParenthese(St : String)  : integer;
Var PosParOuv,PosParFer : integer ;
begin
  result:=0 ;
  While (pos('(',st)>0) and (result=0) do
    begin
    PosParOuv:=pos('(',st) ;
    PosParFer:=pos(')',st) ;
    if (PosParFer>0) then St[PosParFer]:='|' else result:=-9 ;
    if (result=0) and (PosParFer<PosParOuv) then result:=-10 ;
    St[PosParOuv]:='|' ;
    end ;
  if (result=0) then
  if (pos(')',st)>0) then result:=-8 ;
end ;



function TOF_AFREVFORMULE.VerifpasVide(St : string)  :  integer ;
begin
result:=0 ;
if (st='') then result:=-17 ;
end ;

function TOF_AFREVFORMULE.VerifOperateur(St : String)  : integer;
begin
  result:=0 ;
  if st<>'' then
    if St[length(st)] in ['*','+','-','/'] then result:=-11 ;
end ;

function TOF_AFREVFORMULE.SetValIndiceFormule(pStIndice : HString) : variant;
begin
result:=1 ;
end ;

function TOF_AFREVFORMULE.PrepareFormuleAvantInterpretation(St : String)  : String ;
var
  PosVirgule          : Integer;
  PosParentheseFermee : Integer;
  StaJeter            : string;

begin
  // mysterieux
  st:=Stringreplace(st,' ','', [rfReplaceAll,rfIgnoreCase]) ;
  // si il y a un blanc que l'on remplace Gfoemule ne marche plus ;
  While (pos(';',st)>0) do
   begin
   PosVirgule:=pos(';',st) ;
   PosParentheseFermee:=pos('}',st) ;
   StaJeter:=copy(st,PosVirgule,PosParentheseFermee-PosVirgule+1)  ;
   St:=Stringreplace(st,StaJeter,'', [rfReplaceAll,rfIgnoreCase]) ;
   end ;
  st:=Stringreplace(st,'ARR{','', [rfReplaceAll,rfIgnoreCase]) ;
  result:=St ;
end ;
 
function TOF_AFREVFORMULE.VerifGformule(St : HString)  : integer;
var s : string ;
begin
result:=0 ;
st:=PrepareFormuleAvantInterpretation(St) ;
st:= '{"0.###,0"'+st+'}';
s := GFormule(St, SetValIndiceFormule, nil, 1);
if s<>'' then
  begin
  if strtofloat(s)<>1 then result:=-19
  end
  else result:=-19 ;
end ;


procedure TOF_AFREVFORMULE.verification(EnSilence : boolean) ;
Var St : String ;
UneErreur: integer ;
begin
  FormuleOk:=false ;
  st:=MemoFormule.Lines.Text ;
  UneErreur:=VerifpasVide(St) ;
  if UneErreur=0 then UneErreur:=VerifContenuCrochets(St);
  if UneErreur=0 then UneErreur:=VerifArrondi(St) ;
  if UneErreur=0 then UneErreur:=VerifParenthese(St) ;
  if UneErreur=0 then UneErreur:=VerifOperateur(St) ;
  if UneErreur=0 then UneErreur:=VerifIndiceIndice0(uppercase(St)) ;
  if UneErreur=0 then UneErreur:=Verifgformule(uppercase(St)) ;
  if UneErreur<0 then
    begin
//    if not EnSilence then
      PGIBoxAF(TextedesErreurs[-UneErreur], Ecran.Caption);
    end
    Else
    begin
    //Success
    FormuleOk:=True ;
    if not EnSilence then
      PGIInfoAF(TexteMessage[2], Ecran.Caption);
    end

end ;

procedure TOF_AFREVFORMULE.BVerificationClick(sender : Tobject) ;
begin
verification(false) ;
end ;


procedure TOF_AFREVFORMULE.VideChamps ;
begin
  InitialiseLabels ;
  setcontroltext('AFE_FORCODE','');
  setcontroltext('AFE_NBDECCOEF','0');
  setcontroltext('AFE_ARRONDICOEF','0');
  setcontroltext('AFE_NBMOIS','0');
  setcontroltext('AFE_FOREXPRESSION','');
  setcontroltext('AFE_FORLIB','');
  setcontroltext('AFE_FORDESC',''); 
  setcontroltext('AFE_FORTYPE','REV');
  setcontroltext('AFE_INDTYPE','COS');
end ;                                 

procedure TOF_AFREVFORMULE.BinsertClick(sender : Tobject) ;
begin
  VideChamps ;
  Action:=taCreat ;
  MontreBouton;
end ;

procedure TOF_AFREVFORMULE.BDeleteClick(sender : Tobject) ;
begin         
  if PGIAskAF(TexteMessage[1], Ecran.Caption) <> mrYes then Exit;
  ExecuteSQL('DELETE FROM AFFORMULE WHERE AFE_FORCODE="'+TobFormule.Detail[0].GetValue('AFE_FORCODE')+'"') ;
  BinsertClick(Nil) ;
end ;

procedure TOF_AFREVFORMULE.BdefaireClick(sender : Tobject) ;
begin
  ChargeChamps  ;
end ;


function TOF_AFREVFORMULE.PositionneLabels(StLabel  :string ) : boolean;
var i : integer ;
Lb : ThLabel ;
begin
  i:=1   ;
  result:=false ;
  While (i<=10) and (not result) do
    begin
    Lb:=ThLabel(Getcontrol('AFE_INDCODE'+inttostr(i))) ;
    if (Lb.caption='') then
      begin
      Lb.caption:=StLabel ;
      result:=True ;
      end ;
    inc(i) ;
    end ;
end ;

procedure TOF_AFREVFORMULE.InitialiseLabels ;
var i : integer ;
Lb : ThLabel ;
begin
  i:=1   ;
  While (i<=10)   do
    begin
    Lb:=ThLabel(Getcontrol('AFE_INDCODE'+inttostr(i))) ;
    Lb.caption:='' ;
    inc(i) ;
    end ;
end ;

function TOF_AFREVFORMULE.TrouveIndiceDansGrid(Stindice  :string) : boolean;
var i : integer ;
begin
  i:=0 ;
  while ((i<GridIndice.rowCount) and (uppercase(GridIndice.Cells[0,i])<>Stindice)) do inc(i) ;
  if i<GridIndice.rowCount then
    result:=PositionneLabels(GridIndice.Cells[0,i]) else
    result:=false ;
end ;

Function  TOF_AFREVFORMULE.Majindice : boolean  ;
var st,Stindice : String ;
PosOuv,PosFer : integer ;
begin
  InitialiseLabels ;
  st:=MemoFormule.Lines.Text ;
  result:=True ;
  While (pos('[',st)>0) do
    begin
    PosOuv:=pos('[',st) ;
    PosFer:=pos(']',st) ;
    if PosFer>0 then
      begin
      Stindice:=copy(St,PosOuv+1,PosFer-posOuv-1)  ;
      St[PosFer]:='|';
      result:=TrouveIndiceDansGrid(uppercase(Stindice)) ;
      end ;
    St[PosOuv]:='|';
    end ;
FormuleOk:=false ;
SetFocusControl('AFE_FOREXPRESSION') ;
end ;

procedure TOF_AFREVFORMULE.BtnAjouteIndiceClick(sender : Tobject) ;
begin
  MemoFormule.Seltext:='['+Gridindice.Cells[0,Gridindice.Row]+']' ;
  MajIndice ;
end ;

procedure TOF_AFREVFORMULE.BtnAjouteIndice0Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='['+Gridindice.Cells[0,Gridindice.Row]+fin_indice_initial+']' ;
  MajIndice ;
end ;

procedure TOF_AFREVFORMULE.BtnRetrancheIndiceClick(sender : Tobject) ;
Var St : String ;
begin
  St:=MemoFormule.text ;
  St:=StringReplace(St,'['+Gridindice.Cells[0,Gridindice.Row]+']','', [rfReplaceAll,rfIgnoreCase]) ;
  MemoFormule.text :=St ;
  MajIndice ;
end ;

procedure TOF_AFREVFORMULE.BtnRetrancheIndice0Click(sender : Tobject) ;
Var St : String ;
begin
  St:=MemoFormule.text ;
  St:=StringReplace(St,'['+Gridindice.Cells[0,Gridindice.Row]+fin_indice_initial+']' ,'', [rfReplaceAll,rfIgnoreCase]) ;
  MemoFormule.text :=St ;
  MajIndice ;
end ;

procedure TOF_AFREVFORMULE.gerefocus ;
begin
FormuleOk:=false ;
SetFocusControl('AFE_FOREXPRESSION') ;
end ;

procedure TOF_AFREVFORMULE.BTNPAROUVClick(sender : Tobject) ;
begin
  MemoFormule.seltext:='(' ;
  gerefocus ;
end ;

procedure TOF_AFREVFORMULE.BTNPARFERClick(sender : Tobject) ;
begin
  MemoFormule.seltext:=')' ;
  gerefocus ;

end ;

procedure TOF_AFREVFORMULE.BTNARRONDIClick(sender : Tobject) ;
begin
  MemoFormule.Seltext:='ARR{' ;
  gerefocus ;

end ;

procedure TOF_AFREVFORMULE.BTNARRONDIFINClick(sender : Tobject) ;
begin
  MemoFormule.Seltext:='}' ;
  gerefocus ;

end ;

procedure TOF_AFREVFORMULE.Btn1Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='1' ;
   gerefocus ;

end ;

procedure TOF_AFREVFORMULE.Btn2Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='2' ;
   gerefocus ;

end ;

procedure TOF_AFREVFORMULE.Btn3Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='3' ;
    gerefocus ;

end ;

procedure TOF_AFREVFORMULE.Btn4Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='4' ;
  gerefocus ;

end ;

procedure TOF_AFREVFORMULE.Btn5Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='5' ;
   gerefocus ;

end ;

procedure TOF_AFREVFORMULE.Btn6Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='6' ;
    gerefocus ;

  end ;

procedure TOF_AFREVFORMULE.Btn7Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='7' ;
    gerefocus ;

end ;

procedure TOF_AFREVFORMULE.Btn8Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='8' ;
    gerefocus ;

end ;

procedure TOF_AFREVFORMULE.Btn9Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='9' ;
    gerefocus ;

end ;

procedure TOF_AFREVFORMULE.Btn0Click(sender : Tobject) ;
begin
  MemoFormule.Seltext:='0' ;
    gerefocus ;

end ;

procedure TOF_AFREVFORMULE.BTNPOINTVIRGULEClick(sender : Tobject) ;
begin
  MemoFormule.Seltext:=';';
    gerefocus ;

end ;


procedure TOF_AFREVFORMULE.BTNDECIMALSEPARATORClick(sender : Tobject) ;
begin
//  MemoFormule.Seltext:=DecimalSeparator ;
  MemoFormule.Seltext:='.' ; // ca peut varier d'une machine a l'autre
    gerefocus ;

end ;

procedure TOF_AFREVFORMULE.BTNPLUSClick(sender : Tobject) ;
begin
  MemoFormule.Seltext:='+';
    gerefocus ;

end ;

procedure TOF_AFREVFORMULE.BTNMOINSClick(sender : Tobject) ;
begin
  MemoFormule.Seltext:='-';
    gerefocus ;

end ;

procedure TOF_AFREVFORMULE.BTNFOISClick(sender : Tobject) ;
begin
  MemoFormule.Seltext:='*';
    gerefocus ;

end ;

procedure TOF_AFREVFORMULE.BTNDIVClick(sender : Tobject) ;
begin
  MemoFormule.Seltext:='/';
    gerefocus ;

end ;

{procedure TOF_AFREVFORMULE.BEFFClick(sender : Tobject) ;
begin
  MemoFormule.text:='';
    gerefocus ;

end ;
}
 
procedure TOF_AFREVFORMULE.MemoFormuleKeyPress(Sender: TObject; var Key: Char);
begin
 if not (Key in [#3,'(','.',' ',')',';',#8,'0'..'9','*','+','-','/']) then key:=#0 ;
end ;

procedure TOF_AFREVFORMULE.BIndiceClick(sender : Tobject) ;
begin
  AFLanceFiche_MulValIndice;
end;

procedure TOF_AFREVFORMULE.BFormuleEditionClick(sender : Tobject) ;
begin
  AglLanceFicheAFREVFORMULEEDIT('', 'FORCODE='+ GetControlText('AFE_FORCODE') + ';' + 'FORMULE=' + GetControlText('AFE_FOREXPRESSION'));
end;

procedure TOF_AFREVFORMULE.GridIndiceKeyPress(Sender: TObject; var Key: Char);
Var i : integer ;
begin
  i:=0 ;
  FrappeClavier:=Uppercase(FrappeClavier+Key) ;
  //While (i<GridIndice.RowCount) and (pos(FrappeClavier,uppercase(GridIndice.Cells[0,i]))<=0) do i:=i+1 ;
  While (i<GridIndice.RowCount) and (copy(uppercase(GridIndice.Cells[0,i]),1,length(FrappeClavier))<>FrappeClavier) do i:=i+1 ;
  if i>=GridIndice.RowCount then i:=0 ;
  GridIndice.GotoRow(i) ;
  LBFRAPPECLAVIER.caption:=FrappeClavier ;
end ;


procedure TOF_AFREVFORMULE.OnNew ;
begin
  Inherited ;

end ;

procedure TOF_AFREVFORMULE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFREVFORMULE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFREVFORMULE.MontreBouton;
begin
  Bdelete.Enabled:=Action=taModif ;
  Binsert.Enabled:=Action=taModif ;
  Bduplication.Enabled:=Action=taModif ;
    if Action<>taModif then
     SetFocusControl('AFE_FORCODE') ;
end ;

procedure TOF_AFREVFORMULE.FabriqueLesTobs;
Var StQformule : String ;
begin
TobFormule:=TOB.Create('Ma Formule',nil,-1);
  if Action=taCreat then
    begin
    TobInsert:=TOB.Create('AFFORMULE',TobFormule,-1);
    end Else
    begin
    try
      StQformule:='SELECT AFE_FORCODE, AFE_INDTYPE, AFE_NBDECCOEF,AFE_NBMOIS, AFE_ARRONDICOEF,' ;
      StQformule:=StQformule+'AFE_INDCODE10, AFE_INDCODE9, AFE_INDCODE8,AFE_INDCODE7, AFE_INDCODE6, AFE_INDCODE5,' ;
      StQformule:=StQformule+'AFE_INDCODE4, AFE_INDCODE3, AFE_INDCODE2,AFE_INDCODE1, AFE_FORTYPE, AFE_FOREXPRESSION,' ;
      StQformule:=StQformule+'AFE_FORLIB,AFE_FORDESC FROM AFFORMULE '+Clause_where ;
      QFormule:=Nil ;
      QFormule := OpenSQL(StQformule, TRUE);
      TobFormule.LoadDetailDB('AFFORMULE','','',QFormule,false) ;
      ChargeChamps  ;
    finally
    Ferme(QFormule) ;
    end ; // Try
   end ; // if TaCrea
end ;

procedure TOF_AFREVFORMULE.ChargeChamps ;


begin
if TobFormule.Detail.count>0 then
  begin

  setcontroltext('AFE_FORCODE',TobFormule.Detail[0].GetValue('AFE_FORCODE')) ;
  setcontroltext('AFE_INDTYPE',TobFormule.Detail[0].GetValue('AFE_INDTYPE')) ;
  setcontroltext('AFE_FORTYPE',TobFormule.Detail[0].GetValue('AFE_FORTYPE')) ;
  // plus utilisé
  //setcontroltext('AFE_NBDECCOEF',TobFormule.Detail[0].GetValue('AFE_NBDECCOEF')) ;
  setcontroltext('AFE_ARRONDICOEF',TobFormule.Detail[0].GetValue('AFE_ARRONDICOEF')) ;
  setcontroltext('AFE_NBMOIS',TobFormule.Detail[0].GetValue('AFE_NBMOIS')) ;
  setcontroltext('AFE_INDCODE1',TobFormule.Detail[0].GetValue('AFE_INDCODE1')) ;
  setcontroltext('AFE_INDCODE2',TobFormule.Detail[0].GetValue('AFE_INDCODE2')) ;
  setcontroltext('AFE_INDCODE3',TobFormule.Detail[0].GetValue('AFE_INDCODE3')) ;
  setcontroltext('AFE_INDCODE4',TobFormule.Detail[0].GetValue('AFE_INDCODE4')) ;
  setcontroltext('AFE_INDCODE5',TobFormule.Detail[0].GetValue('AFE_INDCODE5')) ;
  setcontroltext('AFE_INDCODE6',TobFormule.Detail[0].GetValue('AFE_INDCODE6')) ;
  setcontroltext('AFE_INDCODE7',TobFormule.Detail[0].GetValue('AFE_INDCODE7')) ;
  setcontroltext('AFE_INDCODE8',TobFormule.Detail[0].GetValue('AFE_INDCODE8')) ;
  setcontroltext('AFE_INDCODE9',TobFormule.Detail[0].GetValue('AFE_INDCODE9')) ;
  setcontroltext('AFE_INDCODE10',TobFormule.Detail[0].GetValue('AFE_INDCODE10')) ;
  setcontroltext('AFE_FORTYPE',TobFormule.Detail[0].GetValue('AFE_FORTYPE')) ;
  setcontroltext('AFE_FOREXPRESSION',TobFormule.Detail[0].GetValue('AFE_FOREXPRESSION')) ;
  FormuleOk:=uppercase(TobFormule.Detail[0].GetValue('AFE_FORMULEOK'))='X' ;
  setcontroltext('AFE_FORLIB',TobFormule.Detail[0].GetValue('AFE_FORLIB')) ;
  setcontroltext('AFE_FORDESC',TobFormule.Detail[0].GetValue('AFE_FORDESC')) ;

  end ;  // if Count
end ;
procedure TOF_AFREVFORMULE.OnLoad ;

begin
  Inherited ;
  MontreBouton ;
  InitialiseLabels ;
  if action = taCreat then VideChamps ;
  FrappeClavier:='';
  OldFrappeClavier:='' ;
  RempliListeIndice ;
  MonTimer:=TTimer.Create(Ecran) ;
  MonTimer.Interval:=500 ;
  MonTimer.OnTimer:=VideFrappeClavier ;
  FabriqueLesTobs ;
//  if ExisteCoef then
//    action := taConsult;
end ;

{function TOF_AFREVFORMULE.ExisteCoef : Boolean;
begin

end;
}

function TOF_AFREVFORMULE.verifChampObligatoire : boolean ;
var erreur : integer ;
begin
interditFermeFenetre :=true;
erreur:=0 ;
if getcontroltext('AFE_INDTYPE')='' then erreur:=14 ;
if getcontroltext('AFE_FORTYPE')='' then erreur:=15 ;
if getcontroltext('AFE_FORCODE')='' then erreur:=18 ;
if getcontroltext('AFE_FOREXPRESSION')='' then
  begin
  erreur:=17 ;
  setfocuscontrol('AFE_FOREXPRESSION') ;
  end ;
if getcontroltext('AFE_INDTYPE')='MOY' then
    if getcontroltext('AFE_NBMOIS')='0'  then erreur:=16 ;
if erreur<>0 then  
  PGIBoxAF(TextedesErreurs[erreur], Ecran.Caption);
interditFermeFenetre :=erreur<>0;
result:=erreur=0;
end ;


procedure TOF_AFREVFORMULE.bValiderClick(sender : Tobject) ;
var
  Erreur,Arrondi,NbMois,coef : integer ;
  maTob : Tob;

begin
  MajIndice ;
  interditFermeFenetre :=false ;
    if verifChampObligatoire then
      begin
      verification(true) ;
//      if Not formuleok then
//        HShowMessage('0;'+Ecran.Caption+';'+TexteMessage[3]+';W;O;O;O;','','') ;
//      coef:=0  ;
//      Arrondi:=0 ;
      Val(GetControltext('AFE_NBDECCOEF'),coef,Erreur)  ;
      Val(GetControltext('AFE_ARRONDICOEF'),Arrondi,Erreur) ;
      Val(GetControltext('AFE_NBMOIS'),NbMois,Erreur) ;
    //  TobFormule.Detail[0].PutValue('AFE_FORMULEOK',FormuleOk) ;
      if FormuleOk then TobFormule.Detail[0].putValue('AFE_FORMULEOK','X') else
        TobFormule.Detail[0].putValue('AFE_FORMULEOK','-') ;
      TobFormule.Detail[0].PutValue('AFE_FORCODE',GetControltext('AFE_FORCODE')) ;
      TobFormule.Detail[0].PutValue('AFE_FORTYPE',GetControltext('AFE_FORTYPE')) ;
      TobFormule.Detail[0].PutValue('AFE_INDTYPE',GetControltext('AFE_INDTYPE')) ;

      TobFormule.Detail[0].PutValue('AFE_NBDECCOEF',inttostr(coef)) ;
      TobFormule.Detail[0].PutValue('AFE_ARRONDICOEF',inttostr(Arrondi)) ;
      TobFormule.Detail[0].PutValue('AFE_NBMOIS',inttostr(NbMois)) ;

      TobFormule.Detail[0].PutValue('AFE_INDCODE1',GetControltext('AFE_INDCODE1')) ;
      TobFormule.Detail[0].PutValue('AFE_INDCODE2',GetControltext('AFE_INDCODE2')) ;
      TobFormule.Detail[0].PutValue('AFE_INDCODE3',GetControltext('AFE_INDCODE3')) ;
      TobFormule.Detail[0].PutValue('AFE_INDCODE4',GetControltext('AFE_INDCODE4')) ;
      TobFormule.Detail[0].PutValue('AFE_INDCODE5',GetControltext('AFE_INDCODE5')) ;
      TobFormule.Detail[0].PutValue('AFE_INDCODE6',GetControltext('AFE_INDCODE6')) ;
      TobFormule.Detail[0].PutValue('AFE_INDCODE7',GetControltext('AFE_INDCODE7')) ;
      TobFormule.Detail[0].PutValue('AFE_INDCODE8',GetControltext('AFE_INDCODE8')) ;
      TobFormule.Detail[0].PutValue('AFE_INDCODE9',GetControltext('AFE_INDCODE9')) ;
      TobFormule.Detail[0].PutValue('AFE_INDCODE10',GetControltext('AFE_INDCODE10')) ;
      TobFormule.Detail[0].PutValue('AFE_FOREXPRESSION',GetControltext('AFE_FOREXPRESSION')) ;
      TobFormule.Detail[0].PutValue('AFE_FORLIB',GetControltext('AFE_FORLIB')) ;
      TobFormule.Detail[0].PutValue('AFE_FORDESC',GetControltext('AFE_FORDESC')) ;

//      if Action=taModif then TobFormule.UpdateDB(False) Else TobFormule.insertDB(nil,false) ;
  //    TobFormule.InsertOrUpdateDB(false) ;

      // on passe par la duplication pour que le blob en duplication soit
      // correctement enregistré
      maTob := tob.create('AFFORMULE', nil, -1);
      try
        maTob.dupliquer(TobFormule, true, true);
        maTob.InsertOrUpdateDB(false) ;
      finally
        maTob.Free;
      end;

      Action:=taModif ;
      MontreBouton;
      end ;
   // end ;
end ;

procedure TOF_AFREVFORMULE.VideFrappeClavier(Sender :Tobject) ;
begin
  if (OldFrappeClavier=FrappeClavier) then FrappeClavier:='';
  OldFrappeClavier:=FrappeClavier;
  LBFRAPPECLAVIER.caption:=FrappeClavier ;
end ;

procedure TOF_AFREVFORMULE.RempliListeIndice ;
Var StRequete : String ;
    TobIndice:TOB;
    MaRequete : TQuery ;
    i : integer ;

begin       
  StRequete:=' SELECT AIN_INDCODE,AIN_INDLIBELLE FROM AFINDICE' ;
  StRequete:=StRequete + ' ORDER BY AIN_INDCODE, AIN_INDLIBELLE'; 
  MaRequete:=Nil ;
  TobIndice:=TOB.Create('',nil,-1);
  i:=0 ;
  try
      MaRequete := OpenSQL(StRequete, TRUE);
      TobIndice.LoadDetailDB('','','',MaRequete,false) ;
      GridIndice.RowCount:=TobIndice.Detail.count ;
      while (i<TobIndice.Detail.count) do
        begin
        GridIndice.Cells[0,i]:=TobIndice.Detail[i].GetValue('AIN_INDCODE')  ;
        GridIndice.Cells[1,i]:=TobIndice.Detail[i].GetValue('AIN_INDLIBELLE')  ;
        i:=i+1 ;
        end ;
      finally
      Ferme(MaRequete);
      TobIndice.Free;
      end ;
end;


procedure TOF_AFREVFORMULE.OnArgument (S : String ) ;
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
        if Champ = 'AFE_FORCODE' then Clause_where :='WHERE AFE_FORCODE="'+Valeur+'"';
        if (Champ ='ACTION') and (valeur = 'CREATION') then Action:=taCreat;
        if (Champ ='ACTION') and (valeur = 'MODIFICATION') then Action:=taModif;
        if (Champ ='ACTION') and (valeur = 'CONSULTATION') then Action:=taConsult;
        END;
     Critere:=(Trim(ReadTokenSt(S)));
  END;

  BtnAjouteIndice:=TToolbarButton97(Getcontrol('BTNAJOUTEINDICE')) ;
  BtnAjouteIndice0:=TToolbarButton97(Getcontrol('BTNAJOUTEINDICE0')) ;
  BtnRetrancheIndice:=TToolbarButton97(Getcontrol('BTNRETRANCHEINDICE')) ;
  BtnRetrancheIndice0:=TToolbarButton97(Getcontrol('BTNRETRANCHEINDICE0')) ;
  BTNPAROUV:=TToolbarButton97(Getcontrol('BTNPAROUV')) ;
  BTNPARFER:=TToolbarButton97(Getcontrol('BTNPARFER')) ;
  BTNARRONDI:=TToolbarButton97(Getcontrol('BTNARRONDI')) ;
  BTNARRONDIFIN:=TToolbarButton97(Getcontrol('BTNARRONDIFIN')) ;
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
  BIndice := TToolbarButton97(Getcontrol('BINDICE')) ;
  BFormuleEdition := TToolbarButton97(Getcontrol('BFORMULEEDITION')) ;

  Binsert.OnClick := BinsertClick;
  BDelete.OnClick := BDeleteClick;
  Bdefaire.OnClick := BdefaireClick;
  BIndice.OnClick := BIndiceClick;
  BFormuleEdition.onClick := BFormuleEditionClick;
 
  BDuplication:=TToolbarButton97(Getcontrol('BDUPLICATION')) ;
  BVerification:=TToolbarButton97(Getcontrol('BVERIFICATION')) ;

  BDuplication.OnClick:=BDuplicationClick ;
  BVerification.OnClick:=BVerificationClick;

  BTNDECIMALSEPARATOR:=TToolbarButton97(Getcontrol('BTNDECIMALSEPARATOR')) ;
  BTNPOINTVIRGULE:=TToolbarButton97(Getcontrol('BTNPOINTVIRGULE')) ;

  BTNPLUS:=TToolbarButton97(Getcontrol('BTNPLUS')) ;
  BTNMOINS:=TToolbarButton97(Getcontrol('BTNMOINS')) ;
  BTNFOIS:=TToolbarButton97(Getcontrol('BTNFOIS')) ;
  BTNDIV:=TToolbarButton97(Getcontrol('BTNDIV')) ;
//  BEFF:=TToolbarButton97(Getcontrol('BEFF')) ;
  bValider:=TToolbarButton97(GetControl('bValider')) ;

  BtnAjouteIndice.OnClick:=BtnAjouteIndiceClick ;
  BtnAjouteIndice0.OnClick:=BtnAjouteIndice0Click ;
  BtnRetrancheIndice.OnClick:=BtnRetrancheIndiceClick ;
  BtnRetrancheIndice0.OnClick:=BtnRetrancheIndice0Click ;
  BTNPAROUV.OnClick:=BTNPAROUVClick ;
//  BEFF.OnClick:=BEFFClick ;
  BTNPARFER.OnClick:=BTNPARFERClick ;
  BTNARRONDI.OnClick:=BTNARRONDIClick ;
  BTNARRONDIFIN.OnClick:=BTNARRONDIFINClick ;
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
  BTNPOINTVIRGULE.OnClick:=BTNPOINTVIRGULEClick ;
  BTNPLUS.OnClick:=BTNPLUSClick ;
  BTNMOINS.OnClick:=BTNMOINSClick ;
  BTNFOIS.OnClick:=BTNFOISClick ;
  BTNDIV.OnClick:=BTNDIVClick ;
  bValider.OnClick:=bValiderClick ;
  LBFRAPPECLAVIER:=THLabel(GetControl('LBFRAPPECLAVIER')) ;
  LBFRAPPECLAVIER.caption:='' ;
  GridIndice:=THGrid(getcontrol('GRIDINDICE'));
  GridIndice.OnKeyPress:=GridIndiceKeyPress;

  //ecran.OnKeyPress:=GridIndiceKeyPress;
  MemoFormule:=Tmemo(getControl('AFE_FOREXPRESSION')) ;
  MemoFormule.OnKeyPress:=MemoFormuleKeyPress ;

  GridIndice.ColWidths[0] := 80;
  GridIndice.ColWidths[1] := 160;

end ;

procedure TOF_AFREVFORMULE.OnClose ;
begin
  Inherited ;
  if interditFermeFenetre then begin
    lastError:=-1 ;
    interditFermeFenetre:=False ;
    end else
    TobFormule.free ;
end ;

procedure TOF_AFREVFORMULE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFREVFORMULE.OnCancel () ;
begin
  Inherited ;
end ;

procedure AFLanceFiche_Formule(cle,Action : string ) ;
begin
  AglLanceFiche ('AFF','AFREVFORMULE','',cle,Action);
end ;

Initialization
  registerclasses ( [ TOF_AFREVFORMULE ] ) ;
end.
