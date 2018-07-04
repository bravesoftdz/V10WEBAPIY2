unit TofVerCptaBor;

interface

uses Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
     Mask, StdCtrls, Hctrls, Hcompte, Buttons, ExtCtrls, Ent1, HEnt1, UTILEDT,
     DB, DBTables, HQry, CpteUtil, hmsgbox, SaisUtil, HStatus, RapSuppr,RappType,
     ComCtrls, CRITEDT, HSysMenu,ParamDat, CpteSav, UtilSais,ed_tools,
     Utof,Utob,FE_Main,HRichEdt, HRichOLE,SaisBor;

type TCtrlBor = Record
     Soc,Etab,Exo,Jal,Qualif,Devise,Nature,EcrAn,SaiEu,Axe,General,Auxi,TypeMvt : string ;
     DateComptable,DateDevise : TDateTime ;
     CumulPiece,CumulBor,CumulPieceE,CumulBorE,CumulPieceD,CumulBorD,CumulPct,TxDev,TotEcr,TotEur,TotDev : double ;
     Bor,Piece : integer ;
     end ;

type TErrorInfos = class
     Jal,Refint,Piece,Ligne : string ;
     Date : TDateTime ;
     MsgErr : string ;
     end ;

type TOF_VerCptaBor2 = class(TOF)
     private
       FListe : THGrid ;
       TErreur : TStringList ;
       procedure AfficheListe ;
       procedure OnZoomClick(Sender: TObject);
     public
       procedure OnLoad ; override ;
     end ;

type TOF_VerCptaBor = class(TOF)
     private
       Err_Ligne,Err_Piece,Err_Bor,Err_Ana,Err_Ven,Err_EcrEtVen,LeNombreErreur : integer ;
       Lib,MonnaiePivot,Sortie,TraitEnCours : boolean ;
       TheTob : Tob ;
       DateDeb,DateFin,NPieceDeb,NpieceFin : THEdit ;
       Exo,Etab,TypeEcr,JalDeb,JalFin : THValComboBox ;
       LigGen,LigAna,EquGen,EquAna: TCheckBox;
       VCrit : TCritEdt ;
       MsgBox1,MsgBox2 : THMsgBox ;
       TCritere,TErreur : TStringList ;
       BlocNote,BlocErr : THRichEditOLE ;
       procedure CreateTheTob ;
       procedure AddTheTob(Table : string ; Sql : string) ;
       function  CatchTheTob(QuelTob : string ) : Tob;
       function  RecupTheTob(QuelTob,QuelChamp :string ; NomChamp : Array of String ; ValeurChamp : Array of Variant ; MultiNiveau : Boolean ) : variant;
       function  ExisteTheTob(QuelTob :string ; NomChamp : Array of String ; ValeurChamp : Array of Variant ; MultiNiveau : Boolean ) : boolean ;
       procedure DestroyTheTob ;
       procedure OnChangeExo(Sender : TObject) ;
       procedure OnClickBStop(Sender : TObject) ;
       procedure OnTropErreurs ;
       procedure InitMsgBox1 ;
       procedure InitMsgBox2 ;
       function Erreur1(Err : integer) : integer;
       procedure Erreur2(Prefix : string;T1 : tob;Err : Integer) ;
       procedure ExportErreur ;
       procedure RecupCriteres(var Err : integer) ;
       function RecupMonnaieAutre : string ;
       procedure InitForm(var Err : integer)  ;
       procedure ChargeJournaux ;
       procedure BuildWhere(var Where : string;Prefix : string) ;
       function CreerTobAnaPArtiel(TobAnaTotal,TobEcr : Tob) : Tob ;
       procedure MajCtrlBor(Pref : string ;var CtrlBor : TCtrlBor;T1 : Tob;RazCumulBor,RazCumulPiece : Boolean) ;
       procedure AjoutCumulCtrlBor(Pref : string;var CtrlBor : TCtrlBor;T1 : Tob) ;
       procedure LanceVerif ;
       procedure LanceVerifGen(TobEcr,TobAna : tob) ;
       procedure LanceVerifAna(TobAnaTotal,TobEcr : tob) ;
       function ChangeBordereau(T1 : Tob;CtrlBor : TCtrlBor) : boolean ;
       function ChangePiece(T1 : Tob;CtrlBor : TCtrlBor) : boolean ;
       function ChangeVentil(T1 : Tob;CtrlBor : TCtrlBor) : boolean ;
       function ControlExisteChamps(T1 : Tob;Prefix : string) : integer ;
       function ControlIntegriteGen(T1 : Tob) : integer ;
       function ControlIntegriteAna(T1 : Tob) : integer ;
       function ControlIntegritePie(T1 : Tob;CtrlBor : TCtrlBor) : integer ;
       function ControlIntegriteBor(T1 : Tob;CtrlBor : TCtrlBor) : integer ;
       function ControlIntegriteVen(T1 : Tob;CtrlBor : TCtrlBor) : integer ;
       function EcritureEtVentilation(TVen,TEcr : Tob) : integer ;
       function ControlCumulOk(CtrlBor : TCtrlBor;Bor,Piece : Boolean) : integer ;
       function ControlCumulAnaOk(CtrlBor : TCtrlBor;TEcr : Tob) : integer ;
       function GeneralEtAutreOK(T1 : Tob;Prefix : string) : integer ;
       function EcrAnouveauOK(T1 : Tob;Prefix : string) : Boolean ;
       function DeviseOK(T1 : Tob;Prefix : string) : integer ;
       function MontantAnaOK(T1 : Tob) : integer ;       
       function NatureGenEtNatureAuxOK(NatureGen,NatureAux : string) : boolean ;
     public
       procedure OnLoad ;   override ;
       procedure OnUpdate ; override ;
       procedure OnClose ;  override ;
     end;

implementation

uses
    {$IFDEF eAGLCLIENT}
    MenuOLX
    {$ELSE}
    MenuOLG
    {$ENDIF eAGLCLIENT}
    ;


var TheData : TObject ;

function BuildTob(Table,Champs,Where,Orderby,Groupby : string) : Tob;
var TheTob : Tob; Sql : String ; Q : TQuery ;
begin
if (Champs='') or (Table='') then begin Result:=nil ; Exit ; end ;
Sql:='SELECT '+Champs+' FROM '+Table ;
if Where<>''   then Sql:=Sql+' WHERE '+Where ;
if GroupBy<>'' then Sql:=Sql+' GROUP BY '+GroupBy ;
if OrderBy<>'' then Sql:=Sql+' ORDER BY '+OrderBy ;
Q:=OpenSql(Sql,True) ;
TheTob:=Tob.Create('$'+Table,nil,-1) ;
TheTob.LoadDetailDB(Table,'','',Q,True,True) ;
ferme(Q) ;
Result:=TheTob ;
end ;

procedure TOF_VerCptaBor.OnLoad ;
var Err : integer ;
begin
MsgBox1:=THMsgBox.create(FMenuG);
MsgBox2:=THMsgBox.create(FMenuG);
TCritere:=TStringList.Create ;
TErreur:=TStringList.Create ;
InitMsgBox1 ; InitMsgBox2 ;
CreateTheTob ;
InitForm(Err) ; if Err<0 then begin Erreur1(Err) ; Exit ; end ;
TraitEnCours:=False ;
end ;

procedure TOF_VerCptaBor.OnClose ;
begin
MsgBox1.Free ;
MsgBox2.Free ;
TCritere.Free ;
//VideStringList(TErreur) ;
TErreur.Free ;
DestroyTheTob ;
end ;

procedure TOF_VerCptaBor.OnUpdate ;
var Err : integer ;
begin
LeNombreErreur:=0 ;
RecupCriteres(Err) ; if Err<0 then begin Erreur1(Err); Exit ; end ;
TErreur.Clear ;
BlocNote.Lines.Clear ;
with BlocNote.DefAttributes do begin Color:=clBlue; Style:=[]; end ;
BlocNote.Lines.Add(TraduireMemoire('Vérification en cours ...')) ;
BlocErr.Lines.Clear ;
with BlocErr.SelAttributes do   begin  Color:=clBlue;  Style:=[]; end ;
BlocErr.Lines.Add(TraduireMemoire('Nombre d''erreurs identifiées : 0')) ;
Sortie:=False ; TraitEnCours:=True ;
LanceVerif ;
TraitEnCours:=False ;
end;

procedure TOF_VerCptaBor.OnClickBStop(Sender : TObject) ;
begin
if TraitEnCours=False then Exit ;
if Erreur1(19)=mrYes then Sortie:=True ;
end ;

procedure TOF_VerCptaBor.OnTropErreurs ;
begin
if TraitEnCours=False then Exit ;
if Erreur1(18)=mrYes then Sortie:=True ;
end ;

procedure TOF_VerCptaBor.OnChangeExo(Sender : TObject) ;
var DateTmp : THedit ;
begin
if (Exo.ItemIndex=0) and (Exo.Items.Count-1>0) then
  begin
  DateTmp:=ThEdit.Create(Application) ;
  ExoToDates(Exo.values[1],DateDeb,DateTmp) ;
  ExoToDates(Exo.values[Exo.Items.Count-1],DateTmp,DateFin) ;
  DateTmp.Free ;
  end
  else
  ExoToDates(Exo.value,DateDeb,DateFin) ;
end ;


procedure TOF_VerCptaBor.RecupCriteres(var Err : integer) ;
var EditTmp : THEdit ; CombTmp : THValComboBox ; ChecTmp : TCheckBox ; TExo : TExoDate ;
begin
Fillchar(VCrit,SizeOf(VCrit),#0) ;
if TCritere=nil then begin Err:=17 ; exit ; end ;
if JalDeb.Value>JalFin.Value then begin Err:=20 ; JalDeb.SetFocus ; Exit ; end ;
QuelDateDeExo(Exo.Value,TExo) ;
if TExo.Code<>'' then
  begin
  if StrToDate(DateDeb.Text)<TExo.Deb then begin Err:=22 ; DateDeb.SetFocus ; exit ; end ;
  if StrToDate(DateFin.Text)>TExo.Fin then begin Err:=22 ; DateFin.SetFocus ; exit ; end ;
  TCritere.Add('EXERCICE>="'+TExo.Code+'"');
  end ;
if DateDeb.Text>DateFin.Text then begin Err:=21 ; DateDeb.SetFocus ; exit ; end ;
if NPieceDeb.Text>NPieceFin.Text then begin Err:=23 ; NPieceDeb.SetFocus ; exit ; end ;
// Ajout des critères dans la stringlist
if TExo.Code<>''      then TCritere.Add('EXERCICE="'+TExo.Code+'"') ;
//if DateDeb.Text<>''   then TCritere.Add('DATECOMPTABLE>="'+UsDateTime(StrToDate(DateDeb.Text))+'"') ;
//if DateFin.Text<>''   then TCritere.Add('DATECOMPTABLE<="'+UsDateTime(StrToDate(DateFin.Text))+'"') ;
if Etab.Value<>''     then TCritere.Add('ETABLISSEMENT="'+Etab.Value+'"') ;
if TypeEcr.Value<>''  then TCritere.Add('QUALIFPIECE="'+TypeEcr.Value+'"') ;
if NPieceDeb.Text<>'' then TCritere.Add('NUMEROPIECE>="'+NPieceDeb.Text+'"') ;
if NPieceFin.Text<>'' then TCritere.Add('NUMEROPIECE<="'+NPieceFin.Text+'"') ;
Err:=0 ;
end ;

procedure TOF_VerCptaBor.InitForm(var Err : integer) ;
var Button : TButton ;
begin
//init des controls
Err:=1 ;  JalDeb:=THValComboBox(GetControl('JALDEB')) ;     if JalDeb=nil     then Exit ;
Err:=2 ;  JalFin:=THValComboBox(GetControl('JALFIN')) ;     if JalFin=nil     then Exit ;
Err:=3 ;  Exo:=THValComboBox(GetControl('EXO')) ;           if Exo=nil        then Exit else Exo.ItemIndex:=0 ;
Err:=4 ;  DateDeb:=THEdit(GetControl('DATEDEB')) ;          if DateDeb=nil    then Exit else DateDeb.Text:=StDate1900;
Err:=5 ;  DateFin:=THEdit(GetControl('DATEFIN')) ;          if DateFin=nil    then Exit else DateFin.Text:=StDate2099;
Err:=6 ;  Etab:=THValComboBox(GetControl('ETAB')) ;         if Etab=nil       then Exit else Etab.ItemIndex:=0 ;
Err:=7 ;  TypeEcr:=THValComboBox(GetControl('TYPEECR')) ;   if TypeEcr=nil    then Exit else TypeEcr.ItemIndex:=0 ;
Err:=8 ;  NPieceDeb:=THEdit(GetControl('NPIECEDEB')) ;      if NPieceDeb=nil  then Exit else NPieceDeb.Text:='0';
Err:=9 ;  NPieceFin:=THEdit(GetControl('NPIECEFIN')) ;      if NPieceFin=nil  then Exit else NPieceFin.Text:='999999999';
Err:=10 ; LigGen:=TCheckBox(GetControl('LIGGEN')) ;         if LigGen=nil     then Exit else LigGen.Checked:=True ;
Err:=11 ; LigAna:=TCheckBox(GetControl('LIGANA')) ;         if LigAna=nil     then Exit else LigAna.Checked:=True ;
Err:=12 ; EquGen:=TCheckBox(GetControl('EQUGEN')) ;         if EquGen=nil     then Exit else EquGen.Checked:=True ;
Err:=13 ; EquAna:=TCheckBox(GetControl('EQUANA')) ;         if EquAna=nil     then Exit else EquAna.Checked:=True ;
Err:=14 ; BlocNote:=THRichEditOLE(GetControl('BLOCNOTE')) ; if BlocNote=nil   then Exit ;
Err:=15 ; BlocErr:=THRichEditOLE(GetControl('BLOCERR')) ;   if BlocErr=nil    then Exit ;
Err:=16 ; Button:=TButton(GetControl('BSTOP')) ;            if Button=nil     then Exit ;
//init des valeurs des controls
if not assigned(Button.onclick)  then Button.OnCLick:=OnCLickBStop ;
if not assigned(Exo.Onchange)    then Exo.OnChange:=OnChangeExo ;
OnChangeExo(Exo) ;
ChargeJournaux ;
JalDeb.ItemIndex:=0 ;
JalFin.ItemIndex:=JalFin.Items.Count-1 ;
Err:=0 ;
end ;

procedure TOF_VerCptaBor.ChargeJournaux ;
var TobJal,T1 : Tob ; i : integer ;
begin
TobJal:=CatchTheTob('JOURNAL') ; if TobJal=nil then exit ;
for i:=0 to TobJal.Detail.Count-1 do
  begin
  T1:=TobJal.Detail[i] ;
  if (T1.GetValue('J_MODESAISIE')='BOR') or (T1.GetValue('J_MODESAISIE')='LIB') then
    begin
    JalDeb.Values.Add(T1.GetValue('J_JOURNAL')) ;
    JalDeb.Items.Add(T1.GetValue('J_LIBELLE')) ;
    JalFin.Values.Add(T1.GetValue('J_JOURNAL')) ;
    JalFin.Items.Add(T1.GetValue('J_LIBELLE')) ;
    end ;
  end;
end ;

procedure TOF_VerCptaBor.BuildWhere(var Where : string;Prefix : string);
var i : integer ; Okok : boolean ;
begin
if Where<>'' then Okok:=False else Okok:=True ;
i:=TCritere.Count-1;
while i>=0 do
 begin
 if Okok then Okok:=False  else Where:=Where+' AND' ;
 Where:=Where+' '+Prefix+TCritere.strings[i] ;
 Dec(i) ;
 end;
end;

function TOF_VerCptaBor.CreerTobAnaPArtiel(TobAnaTotal,TobEcr : Tob) : Tob ;
var T1,T2,TRes : Tob ;i : integer ;
begin
TRes:=nil ;
for i:=0 to TobAnaTotal.Detail.Count-1 do
  begin
  T1:=TobAnaTotal.Detail[i] ;
  if  (T1.GetValue('Y_NUMEROPIECE')  =TobEcr.GetValue('E_NUMEROPIECE'))
  and (T1.GetValue('Y_NUMLIGNE')     =TobEcr.GetValue('E_NUMLIGNE'))
  and (T1.GetValue('Y_EXERCICE')     =TobEcr.GetValue('E_EXERCICE'))
  and (T1.GetValue('Y_SOCIETE')  =TobEcr.GetValue('E_SOCIETE'))
  and (T1.GetValue('Y_ETABLISSEMENT')=TobEcr.GetValue('E_ETABLISSEMENT'))
  and (T1.GetValue('Y_JOURNAL')      =TobEcr.GetValue('E_JOURNAL'))
  then
    begin
    if TRes=nil then TRes:=Tob.Create('$ANALITYQU',nil,-1) ;
    T2:=Tob.Create('ANALYTIQU',TRes,-1) ;
    T2.Dupliquer(T1,False,True) ;
    end ;
  end ;
Result:=TRes ;
end ;
//******************************
//*                            *
//*    LANCEMENT DES VERIFS    *
//*                            *
//******************************
procedure TOF_VerCptaBor.LanceVerif ;
var TobEcr,TobAna,TobJal : Tob ; Jal,Where,stWhat,stJal : string ;
    i,NLig,NLig2,Periode,PeriodeFin : integer ; LaDate : TDateTime ;
begin
TobEcr:=nil ; TobAna:=nil ; Nlig:=0 ; Nlig2:=0 ;
TobJal:=CatchTheTob('JOURNAL') ; if TobJal=nil then exit ;
stWhat:='Vérification ' ;
for i:=0 to TobJal.Detail.Count-1 do
  begin
  Jal:=TobJal.Detail[i].GetValue('J_JOURNAL') ;
  stJal:=TobJal.Detail[i].GetValue('J_LIBELLE')+' ('+TobJal.Detail[i].GetValue('J_JOURNAL')+')' ;
  if (Jal<Jaldeb.Value) or (Jal>JalFin.Value) then Continue ;
  if (TobJal.Detail[i].GetValue('J_MODESAISIE')='BOR') or (TobJal.Detail[i].GetValue('J_MODESAISIE')='LIB') then
    begin
    if Nlig<>0 then begin BlocNote.Lines.Delete(Nlig2) ; BlocNote.Lines.Delete(Nlig) ; Nlig2:=0 ; end ;
    NLig:=BlocNote.Lines.Add(TraduireMemoire(stWhat+' du journal '+stJal+' en cours ...')) ;
    Application.ProcessMessages ;
    if TobJal.Detail[i].GetValue('J_MODESAISIE')='LIB' then Lib:=True else Lib:=False ;
    Periode:=0 ; PeriodeFin:=GetPeriode(StrToDate(DateFin.Text)) ;
    LaDate:=StrToDate(DateDeb.Text) ;
    if Nlig2<>0 then BlocNote.Lines.Delete(Nlig2) ;
    NLig2:=BlocNote.Lines.Add(TraduireMemoire(stWhat+'de la periode N° '+IntToStr(Periode))) ;
    while Periode<PeriodeFin do
      begin
      Periode:=GetPeriode(LaDate) ;
      BlocNote.Lines[NLig2]:=TraduireMemoire(stWhat+'de la periode N° '+IntToStr(Periode)) ;
      Application.ProcessMessages ;
      if LigGen.Checked or EquGen.Checked then
        begin
//        Where:='E_JOURNAL="'+Jal+'" AND E_PERIODE='+IntToStr(Periode) ;
        Where:='E_JOURNAL="'+Jal+'" AND E_DATECOMPTABLE>="'+UsDateTime(DebutDeMois(LaDate))
                                +'" AND E_DATECOMPTABLE<="'+UsDateTime(FinDeMois(LaDate))+'"' ;
        BuildWhere(Where,'E_');
        TobEcr:=BuildTob('ECRITURE','*',Where,'E_NUMEROPIECE,E_NUMGROUPEECR,E_NUMLIGNE','') ;
        end ;
      if LigAna.Checked or EquAna.Checked then
        begin
        Where:='Y_JOURNAL="'+Jal+'" AND Y_PERIODE='+IntToStr(Periode) ;
        BuildWhere(Where,'Y_');
        TobAna:=BuildTob('ANALYTIQ','*',Where,'Y_NUMEROPIECE,Y_NUMLIGNE','') ;
        end ;
      if TobEcr<>nil then LanceVerifGen(TobEcr,TobAna)
                     else LanceVerifAna(TobAna,nil) ;
      TobEcr.free ; TobEcr:=nil ;
      TobAna.free ; TobAna:=nil ;
      LaDate:=FinDeMois(LaDate)+1 ;
      if Sortie then break ;
      end ;
      BlocNote.Lines[NLig2]:=BlocNote.Lines[NLig2]+' Terminée' ;
      Application.ProcessMessages ;
    end ;
  if sortie then break ;
  end ;
if LeNombreErreur>0 then ExportErreur ;
end;

procedure TOF_VerCptaBor.LanceVerifGen(TobEcr,TobAna : Tob) ;
var Err,i : integer ; CtrlBor : TCtrlBor ; T1 : Tob ;
begin
T1:=nil ;
if not LigGen.Checked and not EquGen.Checked then exit ;
if TobEcr.Detail.Count<=0 then Exit ;
MajCtrlBor('E_',CtrlBor,TobEcr.Detail[0],True,True) ;
for i:=0 to TobEcr.Detail.Count-1 do
  begin
  T1:=TobEcr.Detail[i] ;
  if LigGen.Checked then begin Err:=ControlIntegriteGen(T1) ; if Err<>0 then Erreur2('E_',T1,Err) ; end ;
  if EquGen.Checked then
    begin
    if ChangeBordereau(T1,CtrlBor) then
      begin
      Err:=ControlCumulOK(CtrlBor,True,True) ;if Err<>0 then Erreur2('E_',TobEcr.Detail[i-1],Err) ;
      MajCtrlBor('E_',CtrlBor,T1,True,True) ;
      end
      else
      begin
      if ChangePiece(T1,CtrlBor) then
        begin
        Err:=ControlCumulOK(CtrlBor,False,True) ; if Err<>0 then Erreur2('E_',TobEcr.Detail[i-1],Err) ;
        Err:=ControlIntegriteBor(T1,CtrlBor) ;    if Err<>0 then Erreur2('E_',T1,Err) ;
        MajCtrlBor('E_',CtrlBor,T1,False,True) ;
        end
        else
        begin
        Err:=ControlIntegritePie(T1,CtrlBor) ; if Err<>0  then Erreur2('E_',T1,Err) ;
        end ;
      end ;
      if T1.GetValue('E_ANA')='X' then LanceVerifAna(TobAna,T1) ;
      AjoutCumulCtrlBor('E_',CtrlBor,T1) ;
    end ;
  if Sortie then break ;
  end ;
if EquGen.Checked and (T1<>nil)then
  begin
  Err:=ControlIntegriteBor(T1,CtrlBor) ;   if Err<>0 then begin Erreur2('E_',T1,Err) ; end ;
  Err:=ControlCumulOK(CtrlBor,True,True) ; if Err<>0 then begin Erreur2('E_',T1,Err) ; end ;
  end ;
end ;

procedure TOF_VerCptaBor.LanceVerifAna(TobAnaTotal,TobEcr : Tob) ;
var Err,i : integer ; CtrlAna : TCtrlBor ; T1,TobAna : Tob ;
begin
T1:=nil ;
if not LigAna.Checked and not EquAna.Checked then exit ;
if TobEcr=nil then TobAna:=TobAnaTotal else TobAna:=CreerTobAnaPartiel(TobAnaTotal,TobEcr) ;
if TobAna=nil then Exit ;
if TobAna.Detail.Count>0 then
  begin
  MajCtrlBor('Y_',CtrlAna,TobAna.Detail[0],True,True) ;
  for i:=0 to TobAna.Detail.Count-1 do
    begin
    T1:=TobAna.Detail[i] ;
    if LigAna.Checked then begin Err:=ControlIntegriteAna(T1) ;  if Err<>0 then Erreur2('Y_',T1,Err) ; end ;
    if EquAna.Checked then
      begin
      if ChangeVentil(T1,CtrlAna) then
        begin
        Err:=ControlCumulAnaOK(CtrlAna,TobEcr) ;if Err<>0 then Erreur2('Y_',TobAna.Detail[i-1],Err) ;
        MajCtrlBor('Y_',CtrlAna,T1,True,True) ;
        end
        else
        begin
        Err:=ControlIntegriteVen(T1,CtrlAna) ; if Err<>0 then Erreur2('Y_',T1,Err) ;
        end ;
      AjoutCumulCtrlBor('Y_',CtrlAna,T1) ;
      end ;
    if TobEcr<>nil then begin Err:=EcritureEtVentilation(T1,TobEcr) ; if Err<>0 then Erreur2('Y_',T1,Err) ; end ;
    if Sortie then Break ;
    end ;
  if EquAna.Checked then
    begin
    Err:=ControlCumulAnaOK(CtrlAna,TobEcr) ; if Err<>0 then Erreur2('Y_',T1,Err) ;
    end ;
  end;
if TobEcr<>nil then TobAna.Free ;
end ;
//******************************
//*                            *
//*         CHANGEMENTS        *
//*                            *
//******************************
function TOF_VerCptaBor.ChangeBordereau(T1 : Tob;CtrlBor : TCtrlBor) : boolean ;
var BorTmp : integer ;
begin
BorTmp:=T1.GetValue('E_NUMEROPIECE') ;
if (CtrlBor.Etab<>T1.GetValue('E_ETABLISSEMENT'))
or (CtrlBor.Soc <>T1.GetValue('E_SOCIETE'))
or (CtrlBor.Exo <>T1.GetValue('E_EXERCICE'))
or (CtrlBor.Bor <>BorTmp)
then Result:=True
else Result:=False ;
end ;

function TOF_VerCptaBor.ChangePiece(T1 : Tob;CtrlBor : TCtrlBor) : boolean ;
var BorTmp,PieTmp : integer ;
begin
PieTmp:=T1.GetValue('E_NUMGROUPEECR') ;
if (ChangeBordereau(T1,CtrlBor))
or (CtrlBor.Piece <>PieTmp)
then Result:=True
else Result:=False ;
end ;

function TOF_VerCptaBor.ChangeVentil(T1 : Tob;CtrlBor : TCtrlBor) : boolean ;
var VentTmp,VentTm2 : integer ;
begin
VentTmp:=T1.GetValue('Y_NUMEROPIECE') ;
VentTm2:=T1.GetValue('Y_NUMLIGNE') ;
if (CtrlBor.Etab<>T1.GetValue('Y_ETABLISSEMENT'))
or (CtrlBor.Soc <>T1.GetValue('Y_SOCIETE'))
or (CtrlBor.Exo <>T1.GetValue('Y_EXERCICE'))
or (CtrlBor.Axe <>T1.GetValue('Y_AXE'))
or (CtrlBor.Jal <>T1.GetValue('Y_JOURNAL'))
or (CtrlBor.Bor  <>VentTmp)
or (CtrlBor.Piece<>VentTm2)
then Result:=True
else Result:=False ;
end ;
//******************************
//*                            *
//*    CONTROLES D'INTEGRITE   *
//*                            *
//******************************
function TOF_VerCptaBor.ControlIntegriteGen(T1 : Tob) : integer ;
var Pref,Dev : string ;ValTmp,ValTmpDev : double ;  Err : integer ;
begin
Pref:='E_' ;
Err:=ControlExisteChamps(T1,Pref) ; if Err<>0 then begin Result:=Err_Ligne+Err ; Exit ; end ;
if T1.GetValue(Pref+'AUXILIAIRE')<>'' then
  if not ExisteTheTob('TIERS',['T_AUXILIAIRE'],[T1.GetValue(Pref+'AUXILIAIRE')],True) then begin Result:=Err_Ligne+8 ; Exit ; end ;
Err:=GeneralEtAutreOK(T1,Pref) ; if Err<>0 then begin Result:=Err_Ligne+Err ; Exit ; end ;
if T1.GetValue(Pref+'TVAENCAISSEMENT')='X' then
  if not ExisteTheTob('REGIMTVA',['CO_CODE'],[T1.GetValue(Pref+'REGIMTVA')],True) then begin Result:=Err_Ligne+16 ; Exit ; end ;
if T1.GetValue(Pref+'TVA')<>'' then
  if not ExisteTheTob('TXCPTTVA',['TV_CODETAUX'],[T1.GetValue(Pref+'TVA')],True) then begin Result:=Err_Ligne+17 ; Exit ; end ;
if T1.GetValue(Pref+'TPF')<>'' then
  if not ExisteTheTob('TXCPTTVA',['TV_CODETAUX'],[T1.GetValue(Pref+'TPF')],True) then begin Result:=Err_Ligne+18 ; Exit ; end ;
if T1.GetValue(Pref+'ECHE')='X' then
  begin
  if not ExisteTheTob('MODEPAIE',['MP_MODEPAIE'],[T1.GetValue(Pref+'MODEPAIE')],True) then begin Result:=Err_Ligne+20 ; Exit ; end ;
  if StrToDate(T1.GetValue(Pref+'DATEECHEANCE'))=0 then begin Result:=Err_Ligne+21 ; Exit ; end ;
  if T1.GetValue(Pref+'NUMECHE')<>1 then begin Result:=Err_Ligne+29 ; Exit ; end ;
  end
  else
  begin
  if T1.GetValue(Pref+'MODEPAIE')<>'' then begin Result:=Err_Ligne+20 ; Exit ; end ;
  if T1.GetValue(Pref+'NUMECHE')<>0 then begin Result:=Err_Ligne+29 ; Exit ; end ;
  end ;
result:=0 ;
end ;

function TOF_VerCptaBor.ControlIntegriteAna(T1 : Tob) : integer ;
var Pref : string ; Err : integer ;
begin
Pref:='Y_' ;
Err:=ControlExisteChamps(T1,Pref) ; if Err<>0 then begin Result:=Err_Ligne+Err ; Exit ; end ;
if not ExisteTheTob('AXE',['X_AXE'],[T1.GetValue(Pref+'AXE')],True) then begin Result:=Err_Ana+1 ; Exit ; end ;
if not ExisteTheTob('SECTION',['S_SECTION'],[T1.GetValue(Pref+'SECTION')],True) then begin Result:=Err_Ana+2 ; Exit ; end ;
Err:=MontantAnaOK(T1) ; if Err<>0 then begin Result:=Err_Ana+Err ; Exit ; end ;
result:=0 ;
end ;

function TOF_VerCptaBor.ControlIntegritePie(T1 : Tob;CtrlBor : TCtrlBor) : integer ;
var MonnaiePivot : boolean ;
begin
if CtrlBor.Jal          <>T1.GetValue('E_JOURNAL')                  then begin Result:=Err_Piece+4 ;  Exit ; end ;
MonnaiePivot:=((T1.GetValue('E_DEVISE')=V_PGI.DevisePivot) or ((T1.GetValue('E_DEVISE')='') and (T1.GetValue('E_SAISIEURO')='-'))) ;
if Lib and MonnaiePivot then
  begin
  if CtrlBor.DateComptable<>StrToDate(T1.GetValue('E_DATECOMPTABLE')) then begin Result:=Err_Piece+5 ;  Exit ; end ;
  if CtrlBor.Nature       <>T1.GetValue('E_NATUREPIECE')              then begin Result:=Err_Piece+6 ;  Exit ; end ;
  end ;
if CtrlBor.Qualif       <>T1.GetValue('E_QUALIFPIECE')              then begin Result:=Err_Piece+7 ;  Exit ; end ;
if CtrlBor.Devise       <>T1.GetValue('E_DEVISE')                   then begin Result:=Err_Piece+8 ; Exit ; end ;
if CtrlBor.TxDev        <>T1.GetValue('E_TAUXDEV')                  then begin Result:=Err_Piece+9 ; Exit ; end ;
if CtrlBor.DateDevise   <>StrToDate(T1.GetValue('E_DATETAUXDEV'))   then begin Result:=Err_Piece+10 ; Exit ; end ;
if CtrlBor.Qualif       <>T1.GetValue('E_ECRANOUVEAU')              then begin Result:=Err_Piece+11 ;  Exit ; end ;
if CtrlBor.SaiEu        <>T1.GetValue('E_SAISIEEURO')               then begin Result:=Err_Piece+12 ;  Exit ; end ;
Result:=0 ;
end ;

function TOF_VerCptaBor.ControlIntegriteBor(T1 : Tob;CtrlBor : TCtrlBor) : integer ;
begin
if CtrlBor.Jal<>T1.GetValue('E_JOURNAL')                            then begin Result:=Err_Bor+4 ;  Exit ; end ;
if CtrlBor.Qualif <>T1.GetValue('E_QUALIFPIECE')                    then begin Result:=Err_Bor+5 ;  Exit ; end ;
if CtrlBor.Devise<>T1.GetValue('E_DEVISE')                          then begin Result:=Err_Bor+6 ; Exit ; end ;
if CtrlBor.TxDev<>T1.GetValue('E_TAUXDEV')                          then begin Result:=Err_Bor+7 ; Exit ; end ;
if CtrlBor.DateDevise<>T1.GetValue('E_DATETAUXDEV')                 then begin Result:=Err_Bor+8 ; Exit ; end ;
if CtrlBor.Qualif <>T1.GetValue('E_ECRANOUVEAU')                    then begin Result:=Err_Bor+9 ;  Exit ; end ;
Result:=0 ;
end ;

function TOF_VerCptaBor.ControlIntegriteVen(T1 : Tob;CtrlBor : TCtrlBor) : integer ;
begin
if CtrlBor.General<>T1.GetValue('Y_GENERAL')                        then begin Result:=Err_Ven+4 ;  Exit ; end ;
if CtrlBor.DateComptable<>StrToDate(T1.GetValue('Y_DATECOMPTABLE')) then begin Result:=Err_Ven+5 ;  Exit ; end ;
if CtrlBor.Qualif <>T1.GetValue('Y_QUALIFPIECE')                    then begin Result:=Err_Ven+6 ;  Exit ; end ;
if CtrlBor.Nature <>T1.GetValue('Y_NATUREPIECE')                    then begin Result:=Err_Ven+7 ;  Exit ; end ;
if CtrlBor.Devise <>T1.GetValue('Y_DEVISE')                         then begin Result:=Err_Ven+8 ;  Exit ; end ;
if CtrlBor.TxDev <>T1.GetValue('Y_TAUXDEV')                         then begin Result:=Err_Ven+9 ;  Exit ; end ;
if CtrlBor.DateDevise <>T1.GetValue('Y_DATETAUXDEV')                then begin Result:=Err_Ven+10 ; Exit ; end ;
if CtrlBor.TotEcr<>T1.GetValue('Y_TOTALECRITURE')                   then begin Result:=Err_Ven+11 ;  Exit ; end ;
if CtrlBor.TotDev<>T1.GetValue('Y_TOTALDEVISE')                     then begin Result:=Err_Ven+12 ;  Exit ; end ;
if CtrlBor.TotEur<>T1.GetValue('Y_TOTALEURO')                       then begin Result:=Err_Ven+13 ;  Exit ; end ;
if CtrlBor.TypeMvt<>T1.GetValue('Y_TYPEMVT')                        then begin Result:=Err_Ven+14 ;  Exit ; end ;
if CtrlBor.Qualif <>T1.GetValue('Y_ECRANOUVEAU')                    then begin Result:=Err_Ven+15 ;  Exit ; end ;
if CtrlBor.SaiEu  <>T1.GetValue('Y_SAISIEEURO')                     then begin Result:=Err_Ven+16 ;  Exit ; end ;
if CtrlBor.Auxi   <>T1.GetValue('Y_AUXILIAIRE')                     then begin Result:=Err_Ven+17 ;  Exit ; end ;
Result:=0 ;
end ;
//******************************
//*                            *
//*      CONTROLES DIVERS      *
//*                            *
//******************************
function TOF_VerCptaBor.ControlExisteChamps(T1 : Tob;Prefix : string) : integer ;
var Ret : integer ;
begin
if not ExisteTheTob('SOCIETE',['SO_SOCIETE'],[T1.GetValue(Prefix+'SOCIETE')],True) then begin Result:=14 ; Exit ; end ;
if not ExisteTheTob('ETABLISS',['ET_ETABLISSEMENT'],[T1.GetValue(Prefix+'ETABLISSEMENT')],True)then begin Result:=15 ; Exit ; end ;
if not ExisteTheTob('EXERCICE',['EX_EXERCICE'],[T1.GetValue(Prefix+'EXERCICE')],True) then begin Result:=1 ; Exit ; end ;
if not ExisteTheTob('JOURNAL',['J_JOURNAL'],[T1.GetValue(Prefix+'JOURNAL')],True) then begin Result:=2 ; Exit ; end ;
if not ExisteTheTob('GENERAUX',['G_GENERAL'],[T1.GetValue(Prefix+'GENERAL')],True) then begin Result:=6 ; Exit ; end ;
if not ExisteTheTob('NATUREPIECE',['CO_CODE'],[T1.GetValue(Prefix+'NATUREPIECE')],True) then begin Result:=11 ; Exit ; end ;
if not ExisteTheTob('QUALIFPIECE',['CO_CODE'],[T1.GetValue(Prefix+'QUALIFPIECE')],True) then begin Result:=12 ; Exit ; end ;
if not ExisteTheTob('TYPEMVT',['CO_CODE'],[T1.GetValue(Prefix+'TYPEMVT')],True) then begin Result:=13 ; Exit ; end ;
if T1.GetValue(Prefix+'EXERCICE')<>QUELEXODT(T1.GetValue(Prefix+'DATECOMPTABLE')) then begin Result:=3 ; Exit ; end ;
if T1.GetValue(Prefix+'NUMEROPIECE')<=0 then begin Result:=4 ; Exit ; end ;
if T1.GetValue(Prefix+'NUMLIGNE')<=0 then begin Result:=5 ; Exit ; end ;
if Prefix='E_' then begin if T1.GetValue(Prefix+'NUMGROUPEECR')<=0 then begin Result:=34 ; Exit ; end ; end
               else begin if T1.GetValue(Prefix+'NUMVENTIL')<=0 then begin Result:=35 ; Exit ; end ;    end ;
if T1.GetValue(Prefix+'LIBELLE')='' then begin Result:=10 ; Exit ; end ;
if not EcrAnouveauOK(T1,Prefix) then begin Result:=27 ; Exit ; end ;
Ret:=DeviseOk(T1,Prefix) ; if Ret<>0 then begin Result:=Ret ; Exit ; end ;
if GetPeriode(T1.GetValue(Prefix+'DATECOMPTABLE'))<>T1.GetValue(Prefix+'PERIODE') then begin Result:=32 ; Exit ; end ;
if NumSemaine(T1.GetValue(Prefix+'DATECOMPTABLE'))<>T1.GetValue(Prefix+'SEMAINE') then begin Result:=33 ; Exit ; end ;
Result:=0 ;
end ;

function TOF_VerCptaBor.EcritureEtVentilation(TVen,TEcr : Tob) : integer ;
var MttAna,MttEcr : double ;
begin
if TVen.GetValue('Y_JOURNAL')<>TEcr.GetValue('E_JOURNAL')                 then begin Result:=Err_EcrEtVen+1 ;  Exit ; end ;
if TVen.GetValue('Y_GENERAL')<>TEcr.GetValue('E_GENERAL')                 then begin Result:=Err_EcrEtVen+2 ;  Exit ; end ;
if TVen.GetValue('Y_DATECOMPTABLE')<>TEcr.GetValue('E_DATECOMPTABLE')     then begin Result:=Err_EcrEtVen+3 ;  Exit ; end ;
if TVen.GetValue('Y_NUMEROPIECE')<>TEcr.GetValue('E_NUMEROPIECE')         then begin Result:=Err_EcrEtVen+4 ;  Exit ; end ;
if TVen.GetValue('Y_NUMLIGNE')<>TEcr.GetValue('E_NUMLIGNE')               then begin Result:=Err_EcrEtVen+5 ;  Exit ; end ;
if TVen.GetValue('Y_EXERCICE')<>TEcr.GetValue('E_EXERCICE')               then begin Result:=Err_EcrEtVen+6 ;  Exit ; end ;
if TVen.GetValue('Y_NATUREPIECE')<>TEcr.GetValue('E_NATUREPIECE')         then begin Result:=Err_EcrEtVen+7 ;  Exit ; end ;
if TVen.GetValue('Y_QUALIFPIECE')<>TEcr.GetValue('E_QUALIFPIECE')         then begin Result:=Err_EcrEtVen+8 ;  Exit ; end ;
if TVen.GetValue('Y_SOCIETE')<>TEcr.GetValue('E_SOCIETE')                 then begin Result:=Err_EcrEtVen+9 ;  Exit ; end ;
if TVen.GetValue('Y_ETABLISSEMENT')<>TEcr.GetValue('E_ETABLISSEMENT')     then begin Result:=Err_EcrEtVen+10;  Exit ; end ;
if TVen.GetValue('Y_DEVISE')<>TEcr.GetValue('E_DEVISE')                   then begin Result:=Err_EcrEtVen+11;  Exit ; end ;
if TVen.GetValue('Y_TAUXDEV')<>TEcr.GetValue('E_TAUXDEV')                 then begin Result:=Err_EcrEtVen+12;  Exit ; end ;
if TVen.GetValue('Y_DATETAUXDEV')<>TEcr.GetValue('E_DATETAUXDEV')         then begin Result:=Err_EcrEtVen+13;  Exit ; end ;
if TVen.GetValue('Y_ECRANOUVEAU')<>TEcr.GetValue('E_ECRANOUVEAU')         then begin Result:=Err_EcrEtVen+14;  Exit ; end ;
if TVen.GetValue('Y_SAISIEEURO')<>TEcr.GetValue('E_SAISIEEURO')           then begin Result:=Err_EcrEtVen+15;  Exit ; end ;
if TVen.GetValue('Y_AUXILIAIRE')<>TEcr.GetValue('E_AUXILIAIRE')           then begin Result:=Err_EcrEtVen+16;  Exit ; end ;
if TVen.GetValue('Y_PERIODE')<>TEcr.GetValue('E_PERIODE')                 then begin Result:=Err_EcrEtVen+17;  Exit ; end ;
if TVen.GetValue('Y_SEMAINE')<>TEcr.GetValue('E_SEMAINE')                 then begin Result:=Err_EcrEtVen+18;  Exit ; end ;
MttAna:=arrondi(TVen.GetValue('Y_TOTALECRITURE'),V_PGI.OkDecP) ;
MttEcr:=arrondi(TEcr.GetValue('E_DEBIT')+TEcr.GetValue('E_CREDIT'),V_PGI.OkDecP) ;
if MttAna<>MttEcr                                                         then begin Result:=Err_EcrEtVen+19;  Exit ; end ;
MttAna:=arrondi(TVen.GetValue('Y_TOTALEURO'),V_PGI.OkDecP) ;
MttEcr:=arrondi(TEcr.GetValue('E_DEBITEURO')+TEcr.GetValue('E_CREDITEURO'),V_PGI.OkDecP) ;
if MttAna<>MttEcr                                                         then begin Result:=Err_EcrEtVen+20;  Exit ; end ;
MttAna:=arrondi(TVen.GetValue('Y_TOTALDEVISE'),V_PGI.OkDecP) ;
MttEcr:=arrondi(TEcr.GetValue('E_DEBITDEV')+TEcr.GetValue('E_CREDITDEV'),V_PGI.OkDecP) ;
if MttAna<>MttEcr                                                         then begin Result:=Err_EcrEtVen+21;  Exit ; end ;
Result:=0 ;
end ;

function TOF_VerCptaBor.ControlCumulOk(CtrlBor : TCtrlBor;Bor,Piece : Boolean) : integer ;
begin
if Piece and (CtrlBor.CumulPiece <>0) then begin Result:=Err_Piece+1 ; Exit ; end ;
if Piece and (CtrlBor.CumulPieceE<>0) then begin Result:=Err_Piece+2 ; Exit ; end ;
if Piece and (CtrlBor.CumulPieceD<>0) then begin Result:=Err_Piece+3 ; Exit ; end ;
if Bor   and (CtrlBor.CumulBor   <>0) then begin Result:=Err_Bor+1 ;   Exit ; end ;
if Bor   and (CtrlBor.CumulBorE  <>0) then begin Result:=Err_Bor+2 ;   Exit ; end ;
if Bor   and (CtrlBor.CumulBorD  <>0) then begin Result:=Err_Bor+3 ;   Exit ; end ;
Result:=0 ;
end ;

function TOF_VerCptaBor.ControlCumulAnaOk(CtrlBor : TCtrlBor;TEcr : Tob) : integer ;
var mtt : double ; Pref : string ; Deci : integer ;
begin
Pref:='E_' ;
if Abs(CtrlBor.CumulBor) <>CtrlBor.TotEcr then begin Result:=Err_Ven+1 ;  Exit ; end ;
if Abs(CtrlBor.CumulBorD)<>CtrlBor.TotDev then begin Result:=Err_Ven+3 ;  Exit ; end ;
if Abs(CtrlBor.CumulBorE)<>CtrlBor.TotEur then begin Result:=Err_Ven+2 ;  Exit ; end ;
//if CtrlBor.CumulPct <> 100           then begin Result:=Err_Ven+18 ; Exit ; end ;
Result:=0 ;
end ;

function TOF_VerCptaBor.GeneralEtAutreOK(T1 : Tob;Prefix : string) : integer ;
var General,Nature,NatureGen,NatureAux,Auxiliaire,EtatLettrage,Lettrage : string ;
Analytique,Echeance,Ventilable,Lettrable,Pointable,Collectif : boolean ; i : integer ;
begin
General     :=T1.GetValue(Prefix+'GENERAL') ;
Auxiliaire  :=T1.GetValue(Prefix+'AUXILIAIRE') ;
Nature      :=T1.GetValue(Prefix+'NATUREPIECE') ;
NatureGen   :=RecupTheTob('GENERAUX','G_NATUREGENE',['G_GENERAL'],[General],True) ;
NatureAux   :=RecupTheTob('TIERS','T_NATUREAUXI',['T_AUXILIAIRE'],[Auxiliaire],True) ;
Collectif   :=(RecupTheTob('GENERAUX','G_COLLECTIF',['G_GENERAL'],[General],True)='X') ;
if not NATUREPIECECOMPTEOK(Nature,General) then begin Result:=7 ; Exit ; end ;
if not NatureGenEtNatureAuxOK(NatureGen,NatureAux) then begin Result:=9 ; Exit ; end ;
Analytique:=(T1.GetValue(Prefix+'ANA')='X') ;
Echeance:=(T1.GetValue(Prefix+'ECHE')='X') ;
Lettrable:=(RecupTheTob('GENERAUX','G_LETTRABLE',['G_GENERAL'],[General],True)='X') or (RecupTheTob('TIERS','T_LETTRABLE',['T_AUXILIAIRE'],[Auxiliaire],True)='X') ;
Pointable:=(RecupTheTob('GENERAUX','G_POINTABLE',['G_GENERAL'],[General],True)='X') ;
Ventilable:=False ;
for i:=1 to 5 do Ventilable:=Ventilable or (RecupTheTob('GENERAUX','G_VENTILABLE'+IntToStr(i),['G_GENERAL'],[General],True)='X') ;
if not Ventilable and Analytique then  begin Result:=30 ; Exit ; end ;
if Ventilable and not Analytique then  begin Result:=35 ; Exit ; end ;
EtatLettrage:=T1.GetValue(Prefix+'ETATLETTRAGE') ;
Lettrage:=T1.GetValue(Prefix+'LETTRAGE') ;
if (Lettrable or Pointable) and not Echeance then begin Result:=31 ; Exit ; end ;
if not Lettrable and not Pointable and Echeance then begin Result:=36 ; Exit ; end ;
if Lettrable then
  begin
  if (EtatLettrage<>'AL') and (EtatLettrage<>'PL') and (EtatLettrage<>'TL') then begin Result:=28 ; Exit ; end ;
  end
  else
  begin
  if (EtatLettrage<>'RI') then begin Result:=28 ; Exit ; end ;
  if Lettrage<>'' then begin Result:=19 ; Exit ; end ;
  end ;
result:=0 ;
end ;

function TOF_VerCptaBor.EcrAnouveauOK(T1 : Tob;Prefix : string) : boolean ;
var Journal,Anouveau : string ;
begin
result:=True ;
Anouveau:=T1.GetValue(Prefix+'ECRANOUVEAU') ;
Journal:=RecupTheTob('JOURNAL','J_NATUREJAL',['J_JOURNAL'],[T1.GetValue(Prefix+'JOURNAL')],True) ;
if (Journal<>'ANO') and (Journal<>'CLO') and (Journal<>'ANA') then  begin if Anouveau<>'N' then Result:=False ; end
else  if Journal='ANO' then begin if (Anouveau<>'H') and (Anouveau<>'OAN') then Result:=False ; end
      else  if Journal='ANA' then begin if (Anouveau<>'H') and (Anouveau<>'OAN') then Result:=False ; end
            else  if Journal='CLO' then begin  if Anouveau<>'C' then Result:=False ; end ;
end ;

function TOF_VerCptaBor.NatureGenEtNatureAuxOK(NatureGen,NatureAux : string) : boolean ;
begin
if ((NatureGen='COF') and (NatureAux<>'FOU')) or ((NatureGen<>'COF') and (NatureAux='FOU'))
or ((NatureGen='COC') and (NatureAux<>'CLI')) or ((NatureGen<>'COC') and (NatureAux='CLI')) then
  result:=False
  else
  result:=True ;
end ;

function TOF_VerCptaBor.DeviseOK(T1 : Tob;Prefix : string) : integer ;
var Dev : string ; ValTmp,ValTmpDev : double ;
begin
Dev:=T1.GetValue(Prefix+'DEVISE') ;
if Dev<>'' then if not ExisteTheTob('DEVISE',['D_DEVISE'],[T1.GetValue(Prefix+'DEVISE')],True) then begin Result:=22 ; Exit ; end ;
if (Dev=V_PGI.DevisePivot) or (Dev=V_PGI.DeviseFongible ) or (Dev='') then
  begin
  ValTmp:=T1.GetValue(Prefix+'CREDIT') ; ValTmpDev:=T1.GetValue(Prefix+'CREDITDEV') ;
  if ValTmp<>ValTmpDev then begin Result:=23 ; Exit ; end ;
  ValTmp:=T1.GetValue(Prefix+'DEBIT') ;  ValTmpDev:=T1.GetValue(Prefix+'DEBITDEV') ;
  if ValTmp<>ValTmpDev then begin Result:=24 ; Exit ; end ;
  if T1.GetValue(Prefix+'TAUXDEV')<>1 then begin Result:=25 ; Exit ; end ;
  if Prefix='E_' then if T1.GetValue(Prefix+'COTATION')<>1 then begin Result:=26 ; Exit ; end ;
  end ;
Result:=0 ;
end ;

function TOF_VerCptaBor.MontantAnaOK(T1 : Tob) : integer ;
var Tot,TotEur,TotDev,Mtt,MttDev,MttEur{,Pct} : double ; Deci : integer ;
begin
{Pct:=T1.GetValue('Y_POURCENTAGE') ; if Pct<=0 then begin Result:=3 ; exit ; end ;}
if T1.GetValue('Y_DEVISE')<>'' then Deci:=RecupTheTob('DEVISE','D_DECIMALE',['D_DEVISE'],[T1.GetValue('Y_DEVISE')],True)
                               else Deci:=2 ;
{Tot:=arrondi(T1.GetValue('Y_TOTALECRITURE')*Pct/100,V_PGI.OkDecP) ;
Mtt:=arrondi(T1.GetValue('Y_DEBIT')+T1.GetValue('Y_CREDIT'),V_PGI.OkDecP) ;
if Tot=0 then begin Result:=5 ; exit ; end ;
if Tot<>Mtt then begin Result:=4 ; exit ; end ;
Tot:=arrondi(T1.GetValue('Y_TOTALEURO')*Pct/100,V_PGI.OkDecE) ;
Mtt:=arrondi(T1.GetValue('Y_DEBITEURO')+T1.GetValue('Y_CREDITEURO'),V_PGI.OkDecE) ;
if Tot=0 then begin Result:=7 ; exit ; end ;
if Tot<>Mtt then begin Result:=6 ; exit ; end ;
Tot:=arrondi(T1.GetValue('Y_TOTALDEVISE')*Pct/100,Deci) ;
Mtt:=arrondi(T1.GetValue('Y_DEBITDEV')+T1.GetValue('Y_CREDITDEV'),Deci) ;
if Tot=0 then begin Result:=9 ; exit ; end ;
if Tot<>Mtt then begin Result:=8 ; exit ; end ;}
Result:=0 ;
end ;
//******************************
//*                            *
//*       RECORD CTRLBOR       *
//*                            *
//******************************
procedure TOF_VerCptaBor.MajCtrlBor(Pref : string;var CtrlBor : TCtrlBor;T1 : Tob;RazCumulBor,RazCumulPiece : Boolean) ;
begin
CtrlBor.Soc    :=T1.GetValue(Pref+'SOCIETE');
CtrlBor.Etab   :=T1.GetValue(Pref+'ETABLISSEMENT');
CtrlBor.Exo    :=T1.GetValue(Pref+'EXERCICE') ;
CtrlBor.Jal    :=T1.GetValue(Pref+'JOURNAL') ;
CtrlBor.Bor    :=T1.GetValue(Pref+'NUMEROPIECE');
CtrlBor.Qualif :=T1.GetValue(Pref+'QUALIFPIECE') ;
CtrlBor.Nature :=T1.GetValue(Pref+'NATUREPIECE') ;
CtrlBor.EcrAn  :=T1.GetValue(Pref+'ECRANOUVEAU') ;
CtrlBor.SaiEu  :=T1.GetValue(Pref+'SAISIEEURO') ;
CtrlBor.Devise :=T1.GetValue(Pref+'DEVISE') ;
CtrlBor.TxDev  :=T1.GetValue(Pref+'TAUXDEV') ;
CtrlBor.DateDevise   :=StrToDate(T1.GetValue(Pref+'DATETAUXDEV'));
CtrlBor.DateComptable:=StrToDate(T1.GetValue(Pref+'DATECOMPTABLE'));
if (Pref<>'Y_') then
  begin
  CtrlBor.Piece :=T1.GetValue(Pref+'NUMGROUPEECR')
  end
  else
  begin
  CtrlBor.General :=T1.GetValue(Pref+'GENERAL');
  CtrlBor.Auxi    :=T1.GetValue(Pref+'AUXILIAIRE');
  CtrlBor.TypeMvt :=T1.GetValue(Pref+'TYPEMVT');
  CtrlBor.TxDev   :=T1.GetValue(Pref+'TAUXDEV');
  CtrlBor.TotEcr  :=T1.GetValue(Pref+'TOTALECRITURE');
  CtrlBor.TotEur  :=T1.GetValue(Pref+'TOTALEURO');
  CtrlBor.TotDev  :=T1.GetValue(Pref+'TOTALDEVISE');
  CtrlBor.Piece   :=T1.GetValue(Pref+'NUMLIGNE');
  CtrlBor.Axe     :=T1.GetValue(Pref+'AXE');
  end ;
if RazCumulPiece then begin CtrlBor.CumulPiece:=0 ; CtrlBor.CumulPieceE:=0 ; CtrlBor.CumulPieceD:=0 ; end ;
if RazCumulBor   then begin CtrlBor.CumulBor  :=0 ; CtrlBor.CumulBorE  :=0 ; CtrlBor.CumulBorD  :=0 ; end ;
//CtrlBor.CumulPct:=0 ;
end ;

procedure TOF_VerCptaBor.AjoutCumulCtrlBor(Pref : string;var CtrlBor : TCtrlBor;T1 : Tob) ;
var mtt : double ; Deci : integer ;
begin
mtt:=arrondi(T1.GetValue(Pref+'DEBIT'),V_PGI.OkDecP)-arrondi(T1.GetValue(Pref+'CREDIT'),V_PGI.OkDecP) ;
CtrlBor.CumulPiece:=arrondi(CtrlBor.CumulPiece+mtt,V_PGI.OkDecP);
CtrlBor.CumulBor:=arrondi(CtrlBor.CumulBor+mtt,V_PGI.OkDecP) ;
mtt:=arrondi(T1.GetValue(Pref+'DEBITEURO'),V_PGI.OkDecE)-arrondi(T1.GetValue(Pref+'CREDITEURO'),V_PGI.OkDecE) ;
CtrlBor.CumulPieceE:=arrondi(CtrlBor.CumulPieceE+mtt,V_PGI.OkDecE) ;
CtrlBor.CumulBorE:=arrondi(CtrlBor.CumulBorE+mtt,V_PGI.OkDecE) ;
if T1.GetValue(Pref+'DEVISE')<>'' then Deci:=RecupTheTob('DEVISE','D_DECIMALE',['D_DEVISE'],[T1.GetValue(Pref+'DEVISE')],True)
                                  else Deci:=2 ;
mtt:=arrondi(T1.GetValue(Pref+'DEBITDEV'),Deci)-arrondi(T1.GetValue(Pref+'CREDITDEV'),Deci) ;
CtrlBor.CumulPieceD:=arrondi(CtrlBor.CumulPieceD+mtt,Deci) ;
CtrlBor.CumulBorD:=arrondi(CtrlBor.CumulBorD+mtt,Deci) ;
//if Pref='Y_' then CtrlBor.CumulPct:=arrondi(CtrlBor.CumulPct+T1.GetValue(Pref+'POURCENTAGE'),2) ;
end ;

function TOF_VerCptaBor.RecupMonnaieAutre : string ;
begin
if VH^.TenueEuro then Result:=RecupTheTob('DEVISE','D_LIBELLE',['D_DEVISE'],[V_PGI.DeviseFongible],True)
else Result:='Euro' ;
end ;
//******************************
//*                            *
//*  GESTION DE LA TOB         *
//*                            *
//******************************
procedure TOF_VerCptaBor.CreateTheTob ;
var T1: Tob ;
begin
TheTob:=Tob.Create('_THETOB',nil,-1) ;
AddTheTob('EXERCICE','SELECT EX_EXERCICE FROM EXERCICE') ;
AddTheTob('JOURNAL','SELECT J_JOURNAL,J_NATUREJAL,J_MODESAISIE,J_LIBELLE FROM JOURNAL') ;
AddTheTob('GENERAUX','SELECT G_GENERAL,G_NATUREGENE,G_POINTABLE,G_LETTRABLE,G_VENTILABLE1,G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5 FROM GENERAUX') ;
AddTheTob('TIERS','SELECT T_AUXILIAIRE,T_NATUREAUXI,T_LETTRABLE FROM TIERS') ;
AddTheTob('SOCIETE','SELECT SO_SOCIETE FROM SOCIETE') ;
AddTheTob('ETABLISS','SELECT ET_ETABLISSEMENT FROM ETABLISS') ;
AddTheTob('QUALIFPIECE','SELECT * FROM COMMUN WHERE CO_TYPE="QFP"' );
AddTheTob('NATUREPIECE','SELECT * FROM COMMUN WHERE CO_TYPE="NTP"' );
AddTheTob('TYPEMVT','SELECT * FROM COMMUN WHERE CO_TYPE="TMC"' );
AddTheTob('REGIMTVA','SELECT * FROM COMMUN WHERE CO_TYPE="RTV"' );
AddTheTob('TXCPTTVA','' );
AddTheTob('MODEPAIE','' );
AddTheTob('DEVISE','SELECT D_DEVISE,D_LIBELLE,D_DECIMALE FROM DEVISE' );
AddTheTob('AXE','SELECT X_AXE FROM AXE' );
AddTheTob('SECTION','SELECT S_SECTION FROM SECTION' );
end ;

procedure TOF_VerCptaBor.AddTheTob(Table : string ; Sql : String) ;
var T1 : tob ; Q : TQuery ;
begin
if Sql<>'' then Q:=OpenSql(Sql,True) else Q:=nil ;
T1:=Tob.Create('_'+Table,TheTob,-1) ;
T1.AddChampSup('NOMTABLE',TRUE) ;  T1.PutValue('NOMTABLE',Table) ;
T1.LoadDetailDB(Table,'','',Q,True);
if Q<>nil then ferme(Q) ;
end ;

function TOF_VerCptaBor.CatchTheTob(QuelTob : string ) : Tob;
begin
Result:=nil ;
if TheTob=nil then exit ;
Result:=TheTob.FindFirst(['NOMTABLE'],[QuelTob],False) ;
end;

function TOF_VerCptaBor.RecupTheTob(QuelTob,QuelChamp :string ; NomChamp : Array of String ; ValeurChamp : Array of Variant ; MultiNiveau : Boolean ) : variant;
var T1,T2 : Tob ;
begin
if TheTob=nil then begin Result:='' ; Exit ; end ;
T2:=nil ;
T1:=TheTob.FindFirst(['NOMTABLE'],[QuelTob],False) ;
if T1<>nil then T2:=T1.FindFirst(NomChamp,ValeurChamp,Multiniveau) ;
if T2=nil then begin Result:=#0 ; exit ; end ;
Result:=T2.GetValue(QuelChamp) ;
end;

function TOF_VerCptaBor.ExisteTheTob(QuelTob :string ; NomChamp : Array of String ; ValeurChamp : Array of Variant ; MultiNiveau : Boolean ) : boolean ;
var T1,T2 : Tob ;
begin
if TheTob=nil then begin Result:=False ; Exit ; end ;
T2:=nil ;
T1:=TheTob.FindFirst(['NOMTABLE'],[QuelTob],False) ;
if T1<>nil then T2:=T1.FindFirst(NomChamp,ValeurChamp,Multiniveau) ;
if T2=nil then Result:=False else Result:=True ;
end;

procedure TOF_VerCptaBor.DestroyTheTob ;
begin
if TheTob<>nil then begin TheTob.free ; TheTob:=nil ; end ;
end ;
//******************************
//*                            *
//*  TRAITMENT DES MSG ERREUR  *
//*                            *
//******************************
//Message d'erreur afficher à l'ecran
function TOF_VerCptaBor.Erreur1(Err : integer) : integer ;
begin
Result:=0 ;
if MsgBox1=nil then exit ;
if Err>MsgBox1.Mess.Count-1 then Err:=0 ;
result:=MsgBox1.Execute(Err,'','') ;
end ;

Procedure TOF_VerCptaBor.InitMsgBox1;
var i : integer ;
begin
if MsgBox1=nil then exit ;
// Erreur de 1..19 erreur de conception !
MsgBox1.Mess.Add('0;;;W;O;O;O');
MsgBox1.Mess.Add('1;Erreur de la fenetre;Le control JALDEB est inaccessible;W;O;O;O');
MsgBox1.Mess.Add('2;Erreur de la fenetre;Le control JALFIN est inaccessible;W;O;O;O');
MsgBox1.Mess.Add('3;Erreur de la fenetre;Le control EXO est inaccessible;W;O;O;O');
MsgBox1.Mess.Add('4;Erreur de la fenetre;Le control DATEDEB est inaccessible;W;O;O;O');
MsgBox1.Mess.Add('5;Erreur de la fenetre;Le control DATEFIN est inaccessible;W;O;O;O');
MsgBox1.Mess.Add('6;Erreur de la fenetre;Le control ETAB est inaccessible;W;O;O;O');
MsgBox1.Mess.Add('7;Erreur de la fenetre;Le control TYPEECR est inaccessible;W;O;O;O');
MsgBox1.Mess.Add('8;Erreur de la fenetre;Le control NPIECEDEB est inaccessible;W;O;O;O');
MsgBox1.Mess.Add('9;Erreur de la fenetre;Le control NPIECEFIN est inaccessible;W;O;O;O');
MsgBox1.Mess.Add('10;Erreur de la fenetre;Le control LIGGEN est inaccessible;W;O;O;O');
MsgBox1.Mess.Add('11;Erreur de la fenetre;Le control LIGANA est inaccessible;W;O;O;O');
MsgBox1.Mess.Add('12;Erreur de la fenetre;Le control EQUGEN est inaccessible;W;O;O;O');
MsgBox1.Mess.Add('13;Erreur de la fenetre;Le control EQUANA est inaccessible;W;O;O;O');
MsgBox1.Mess.Add('14;Erreur de la fenetre;Le control BLOCNOTE est inaccessible;W;O;O;O');
MsgBox1.Mess.Add('15;Erreur de la fenetre;Le control BLOCERR est inaccessible;W;O;O;O');
MsgBox1.Mess.Add('16;Erreur de la fenetre;Le control BSTOP n''existe pas;W;O;O;O');
MsgBox1.Mess.Add('17;Erreur de la fenetre;L''objet TCRITERE n''existe pas;W;O;O;O');
MsgBox1.Mess.Add('18;Vérification de la saisie bordereau;Le nombre d''erreurs est superieur à 500, voulez-vous arreter la traitement;Q;YN;N;N');
MsgBox1.Mess.Add('19;Vérification de la saisie bordereau;Voulez-vous arreter la traitement;Q;YN;N;N');
MsgBox1.Mess.Add('20;Contrôle des mouvements (Bordereaux);Le journal de debut est superieur au journal de fin;W;O;O;O');
MsgBox1.Mess.Add('21;Contrôle des mouvements (Bordereaux);La date de debut est superieur a la date de fin;W;O;O;O');
MsgBox1.Mess.Add('22;Contrôle des mouvements (Bordereaux);La date est en dehors de l''exercice;W;O;O;O');
MsgBox1.Mess.Add('23;Contrôle des mouvements (Bordereaux);Le numero de piece de debut est superieur au numero de piece de fin;W;O;O;O');
end ;
//Message d'erreur enregistré dans la string liste
procedure TOF_VerCptaBor.Erreur2(Prefix : string;T1 : tob;Err : Integer) ;
var infos : TErrorInfos ;
begin
if MsgBox2=nil then exit ;
if Err>MsgBox2.Mess.Count-1 then Err:=0 ;
Infos:=TErrorInfos.Create ;
Infos.Jal   :=T1.GetValue(Prefix+'JOURNAL') ;
Infos.Refint:=T1.GetValue(Prefix+'REFINTERNE') ;
Infos.Piece :=T1.GetValue(Prefix+'NUMEROPIECE') ;
Infos.Ligne :=T1.GetValue(Prefix+'NUMLIGNE') ;
Infos.Date  :=StrToDate(T1.GetValue(Prefix+'DATECOMPTABLE')) ;
if Prefix='E_' then Infos.MsgErr:='Ecriture:' else Infos.MsgErr:='Analytique:' ;
Infos.MsgErr:=Infos.MsgErr+MsgBox2.Mess[Err] ;
TErreur.AddObject('',Infos) ;
LeNombreErreur:=LeNombreErreur+1 ;
if LeNombreErreur=500 then OnTropErreurs ;
BlocErr.Lines.Clear ;
with BlocErr.SelAttributes do   begin  Color:=clRed; end ;
BlocErr.Lines.Add(TraduireMemoire('Nombre d''erreurs identifiées : '+intToStr(LeNombreErreur))) ;
Application.ProcessMessages ;
end ;

Procedure TOF_VerCptaBor.InitMsgBox2;
var i : integer ;
begin
if MsgBox2=nil then exit ;
Err_Ligne:=0 ;
{00}MsgBox2.Mess.Add('Erreur inconnue');
{01}MsgBox2.Mess.Add('L''exercice n''existe pas');
{02}MsgBox2.Mess.Add('Le journal n''existe pas');
{03}MsgBox2.Mess.Add('La date comptable est en dehors de l''exercice');
{04}MsgBox2.Mess.Add('le numero de piece est inferieur à 1');
{05}MsgBox2.Mess.Add('Le numero de ligne est inferieur à 1');
{06}MsgBox2.Mess.Add('Le compte général n''existe pas');
{07}MsgBox2.Mess.Add('Le compte général n''est pas compatible avec la nature de piece');
{08}MsgBox2.Mess.Add('Le compte auxiliaire n''existe pas');
{09}MsgBox2.Mess.Add('Le compte auxiliaire est incompatible avec le compte général');
{10}MsgBox2.Mess.Add('Le libelle est vide');
{11}MsgBox2.Mess.Add('La nature de piece n''existe pas');
{12}MsgBox2.Mess.Add('Le qualifiant piece n''existe pas');
{13}MsgBox2.Mess.Add('Le type de mouvement n''existe pas');
{14}MsgBox2.Mess.Add('Le code société n''existe pas');
{15}MsgBox2.Mess.Add('Le code etablissement n''existe pas');
{16}MsgBox2.Mess.Add('Le regime de tva n''existe pas');
{17}MsgBox2.Mess.Add('Le code tva n''existe pas');
{18}MsgBox2.Mess.Add('Le code tpf n''existe pas');
{19}MsgBox2.Mess.Add('La valeur du champ lettrage est incorrecte');
{20}MsgBox2.Mess.Add('Le mode de paiement n''est pas renseigné ou incorrect');
{21}MsgBox2.Mess.Add('La date d''echeance n''est pas renseigné');
{22}MsgBox2.Mess.Add('La devise n''existe pas');
{23}MsgBox2.Mess.Add('Le montant credit devise ne correspond pas avec le montant credit');
{24}MsgBox2.Mess.Add('Le montant debit devise ne correspond pas avec le montant debit');
{25}MsgBox2.Mess.Add('Le taux de la devise est incorrect');
{26}MsgBox2.Mess.Add('La cotation de la devise est incorrect');
{27}MsgBox2.Mess.Add('Le type d''anouveau est incorrect');
{28}MsgBox2.Mess.Add('L''etat de lettrage est incorrect');
{29}MsgBox2.Mess.Add('Le nombre d''echéance est incorrect');
{20}MsgBox2.Mess.Add('L''écriture n''a pas de ventilation analytique');
{31}MsgBox2.Mess.Add('L''écriture n''a pas d''échéance');
{32}MsgBox2.Mess.Add('La valeur de la periode est incorrecte');
{33}MsgBox2.Mess.Add('La valeur de la semaine est incorrecte');
{34}MsgBox2.Mess.Add('Le numero du groupe d''écritures est inferieur à 1 ');
{35}MsgBox2.Mess.Add('Le numero de la ventilation est négatif');
{20}MsgBox2.Mess.Add('L''écriture ne peux pas avoir de ventilation analytique');
{31}MsgBox2.Mess.Add('L''écriture ne peux pas avoir d''échéance');
Err_Piece:=MsgBox2.Mess.Count ;
{00}MsgBox2.Mess.Add('Erreur inconnue');
{01}MsgBox2.Mess.Add('Le cumul de la piece n''est pas nul');
{02}MsgBox2.Mess.Add('Le cumul de la piece en '+RecupMonnaieAutre+' n''est pas nul');
{03}MsgBox2.Mess.Add('Le cumul de la piece en devise n''est pas nul');
{04}MsgBox2.Mess.Add('Le journal est different de celui de la piece');
{05}MsgBox2.Mess.Add('La date comptable est differente de celle de la piece');
{06}MsgBox2.Mess.Add('La nature de piece est differente de celle de la piece');
{07}MsgBox2.Mess.Add('Le qualifiant piece est different de celui de la piece');
{08}MsgBox2.Mess.Add('La devise est differente de celle de la piece');
{09}MsgBox2.Mess.Add('Le taux de la devise est different de celui de la piece');
{10}MsgBox2.Mess.Add('La date du taux devise est differente de celle de la piece');
{11}MsgBox2.Mess.Add('Le type à nouveau est different de celui de la piece');
{12}MsgBox2.Mess.Add('Le type de saisie euro est different de celui de la piece');
Err_Bor:=MsgBox2.Mess.Count ;
{00}MsgBox2.Mess.Add('Erreur inconnue');
{01}MsgBox2.Mess.Add('Le cumul du bordereau n''est pas nul');
{02}MsgBox2.Mess.Add('Le cumul du bordereau en '+RecupMonnaieAutre+' n''est pas nul');
{03}MsgBox2.Mess.Add('Le cumul du bordereau en devise n''est pas nul');
{04}MsgBox2.Mess.Add('Le journal est different de celui du bordereau');
{05}MsgBox2.Mess.Add('La qualifiant est different de celui du bordereau');
{06}MsgBox2.Mess.Add('La devise est differente de celle du bordereau');
{07}MsgBox2.Mess.Add('Le taux de la devise est different de celui du bordereau');
{08}MsgBox2.Mess.Add('La date du taux devise est differente de celle du bordereau');
{09}MsgBox2.Mess.Add('Le type à nouveau est different de celui du bordereau');
Err_Ana:=MsgBox2.Mess.Count ;
{00}MsgBox2.Mess.Add('Erreur inconnue');
{01}MsgBox2.Mess.Add('L''axe analytique n''existe pas');
{02}MsgBox2.Mess.Add('La section analytique n''existe pas');
{03}MsgBox2.Mess.Add('Le pourcentage de ventilation est incorrect');
{04}MsgBox2.Mess.Add('Le montant de la ventilation est different du pourcentage du montant de l''ecriture');
{05}MsgBox2.Mess.Add('Le total de l''écriture est égal à 0');
{06}MsgBox2.Mess.Add('Le montant de la ventilation en euro est different du pourcentage du montant de l''ecriture');
{07}MsgBox2.Mess.Add('Le total de l''écriture en '+RecupMonnaieAutre+' est égal à 0');
{08}MsgBox2.Mess.Add('Le montant de la ventilation en devise est different du pourcentage du montant de l''ecriture');
{09}MsgBox2.Mess.Add('Le total de l''écriture en devise est égal à 0');
Err_Ven:=MsgBox2.Mess.Count ;
{00}MsgBox2.Mess.Add('Erreur inconnue');
{01}MsgBox2.Mess.Add('Le cumul de la ventilation n''est pas égal au montant de l''écriture');
{02}MsgBox2.Mess.Add('Le cumul de la ventilation en '+RecupMonnaieAutre+' n''est pas égal au montant de l''écriture');
{03}MsgBox2.Mess.Add('Le cumul de la ventilation en devise n''est pas égal au montant de l''écriture');
{04}MsgBox2.Mess.Add('Le compte général est différent de celui de la ventilation');
{05}MsgBox2.Mess.Add('La date comptable est différente de celle de la ventilation');
{06}MsgBox2.Mess.Add('Le qualifiant piece est différent de celui de la ventilation');
{07}MsgBox2.Mess.Add('La nature de piece est différente de celle de la ventilation');
{08}MsgBox2.Mess.Add('La devise est différente de celle de la ventilation');
{19}MsgBox2.Mess.Add('Le taux de la devise est différent de celui de la ventilation');
{10}MsgBox2.Mess.Add('La date du taux de la devise est différente de celle de la ventilation');
{11}MsgBox2.Mess.Add('Le montant de l''écriture est différent de celui de la ventilation');
{12}MsgBox2.Mess.Add('Le montant de l''écriture en devise est différent de celui de la ventilation');
{13}MsgBox2.Mess.Add('Le montant de l''écriture en '+RecupMonnaieAutre+' est différent de celui de la ventilation');
{14}MsgBox2.Mess.Add('Le type de mouvement est différent de celui de la ventilation');
{15}MsgBox2.Mess.Add('Le type d''anouveau est différent de celui de la ventilation');
{16}MsgBox2.Mess.Add('Le type de saisie euro est different de celui de la ventilation');
{17}MsgBox2.Mess.Add('Le compte auxiliaire euro est different de celui de la ventilation');
{18}MsgBox2.Mess.Add('La somme des pourcentage est différent de 100%');
Err_EcrEtVen:=MsgBox2.Mess.Count ;
{00}MsgBox2.Mess.Add('Erreur inconnue');
{01}MsgBox2.Mess.Add('Le journal de la ventilation ne correspond pas à celui de l''écriture');
{02}MsgBox2.Mess.Add('Le compte général de la ventilation ne correspond pas à celui de l''écriture');
{03}MsgBox2.Mess.Add('La date comptable de la ventilation ne correspond pas à celle de l''écriture');
{04}MsgBox2.Mess.Add('Le numero de piece de la ventilation ne correspond pas à celui de l''écriture');
{05}MsgBox2.Mess.Add('Le numero de ligne de la ventilation ne correspond pas à celui de l''écriture');
{06}MsgBox2.Mess.Add('L''exercice de la ventilation ne correspond pas à celui de l''écriture');
{07}MsgBox2.Mess.Add('La nature de piece de la ventilation ne correspond pas à celle de l''écriture');
{08}MsgBox2.Mess.Add('Le qualifiant de la ventilation ne correspond pas à celui de l''écriture');
{09}MsgBox2.Mess.Add('La société de la ventilation ne correspond pas à celle de l''écriture');
{10}MsgBox2.Mess.Add('L''établissement de la ventilation ne correspond pas à celui de l''écriture');
{11}MsgBox2.Mess.Add('La devise de la ventilation ne correspond pas à celle de l''écriture');
{12}MsgBox2.Mess.Add('Le taux de la devise de la ventilation ne correspond pas à celui de l''écriture');
{13}MsgBox2.Mess.Add('La date du taux de la devise de la ventilation ne correspond pas à celle de l''écriture');
{14}MsgBox2.Mess.Add('Le type d''anouveau de la ventilation ne correspond pas à celui de l''écriture');
{15}MsgBox2.Mess.Add('Le type de saisie Euro de la ventilation ne correspond pas à celui de l''écriture');
{16}MsgBox2.Mess.Add('L''auxiliaire de la ventilation ne correspond pas à celui de l''écriture');
{17}MsgBox2.Mess.Add('La periode de la ventilation ne correspond pas à celle de l''écriture');
{18}MsgBox2.Mess.Add('La semaine de la ventilation ne correspond pas à celle de l''écriture');
{19}MsgBox2.Mess.Add('Le montant de l''écriture est différent du total écriture de la ventilation');
{20}MsgBox2.Mess.Add('Le montant de l''écriture en devise est différent du total écriture de la ventilation');
{21}MsgBox2.Mess.Add('Le montant de l''écriture en '+RecupMonnaieAutre+' du total écriture de celui de la ventilation');
end ;

procedure TOF_VerCptaBor.ExportErreur ;
var i : integer ; Form : TForm ; FListe : THGrid ;
begin
TheData:=TErreur ;
AGLLanceFiche('CP', 'VERCPTABOR2', '', '', '') ;
end ;
//******************************
//*                            *
//*    TOF POUR LES MESSAGES   *
//*                            *
//******************************
procedure TOF_VerCptaBor2.OnLoad ;
var BButton : TButton ;
begin
TErreur:=TStringList(TheData) ;
FListe:=THGrid(GetControl('FLISTE'));
BButton:=TButton(GetControl('BZOOM')) ;
if (BButton<>nil) and (not assigned(BButton.OnClick)) then BButton.OnClick:=OnZoomClick ;
AfficheListe ;
end ;

procedure TOF_VerCptaBor2.AfficheListe ;
var i : integer ;
begin
if Fliste=nil then Exit ;
for i:=0 to TErreur.Count-1 do
  begin
  Fliste.Cells[0,Fliste.RowCount-1]:=TErrorInfos(TErreur.Objects[i]).Jal ;
  Fliste.Cells[1,Fliste.RowCount-1]:=TErrorInfos(TErreur.Objects[i]).Refint ;
  Fliste.Cells[2,Fliste.RowCount-1]:=TErrorInfos(TErreur.Objects[i]).Piece+'/'+TErrorInfos(TErreur.Objects[i]).Ligne ;
  Fliste.Cells[3,Fliste.RowCount-1]:=DateToStr(TErrorInfos(TErreur.Objects[i]).Date);
  Fliste.Cells[4,Fliste.RowCount-1]:=TErrorInfos(TErreur.Objects[i]).MsgErr ;
  if i>0 then Fliste.RowCount:=Fliste.RowCount+1 ;
  end;
end ;

procedure TOF_VerCptaBor2.OnZoomClick(Sender: TObject);
Var i : integer ;Chaine : string ;    P  : RParFolio ;
begin
if TErreur=nil then Exit ;
FillChar(P, Sizeof(P), #0) ;
P.ParPeriode:=DateToStr(DebutDeMois(StrToDate(Fliste.Cells[3,Fliste.Row]))) ;
P.ParCodeJal:=Fliste.Cells[0,Fliste.Row] ;
Chaine:=Fliste.Cells[2,Fliste.Row] ;
i:=pos('/',Chaine) ;
P.ParNumFolio:=copy(Chaine,1,i-1) ;
P.ParNumLigne:=StrToInt(copy(Chaine,i+1,length(Chaine)-i)) ;
ChargeSaisieFolio(P, taConsult) ;
end;

Initialization
registerclasses([TOF_VerCptaBor]);
registerclasses([TOF_VerCptaBor2]);
end.
