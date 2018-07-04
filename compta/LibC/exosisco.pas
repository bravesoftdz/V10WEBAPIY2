unit EXOSISCO;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, {$IFNDEF DBXPRESS}dbtables, Db, StdCtrls, Hctrls, Hcompte,
  HRegCpte, Mask, DBCtrls, HSysMenu, hmsgbox, Hqry, HTB97, ExtCtrls,
  HPanel, Grids, DBGrids, HDB{$ELSE}uDbxDataSet{$ENDIF}, DB, StdCtrls, Hctrls, Mask, DBCtrls, hmsgbox,
  Buttons, ExtCtrls, Grids, DBGrids, HDB, Ent1, HEnt1, HCompte, HSysMenu,
  Hqry, HRegCpte, HTB97, UiUtil, HPanel, Paramsoc ;

Procedure FicheExerciceSISCO(Quel : String ; Comment : TActionFiche ; Origine: Byte) ;
function ExoOk(AvecMess : Boolean) : Boolean ;

type
  TFExoSISCO = class(TFFicheListe)
    TEX_EXERCICE  : THLabel;
    EX_EXERCICE   : TDBEdit;
    TEX_ABREGE    : THLabel;
    EX_ABREGE     : TDBEdit;
    TEX_LIBELLE   : THLabel;
    EX_LIBELLE    : TDBEdit;
    TEX_DATEDEBUT : THLabel;
    EX_DATEDEBUT  : TDBEdit;
    TEX_DATEFIN   : THLabel;
    EX_DATEFIN    : TDBEdit;
    TaEX_EXERCICE   : TStringField;
    TaEX_LIBELLE    : TStringField;
    TaEX_ABREGE     : TStringField;
    TaEX_DATEDEBUT  : TDateTimeField;
    TaEX_DATEFIN    : TDateTimeField;
    TaEX_ETATCPTA   : TStringField;
    TaEX_ETATBUDGET : TStringField;
    TaEX_ETATADV    : TStringField;
    TaEX_ETATAPPRO  : TStringField;
    TaEX_ETATPROD   : TStringField;
    TaEX_SOCIETE    : TStringField;
    TaEX_VALIDEE    : TStringField;
    TaEX_DATECUM    : TDateTimeField;
    TExercice: THTable;
    TExerciceEX_EXERCICE: TStringField;
    TExerciceEX_DATEDEBUT: TDateTimeField;
    TExerciceEX_DATEFIN: TDateTimeField;
    Confirm: THMsgBox;
    TaEX_BUDJAL: TStringField;
    EX_BUDJAL: THDBCpteEdit;
    TEX_BUDJAL: TLabel;
    TaEX_DATECUMRUB: TDateTimeField;
    TaEX_NATEXO: TStringField;
    TaEX_DATECUMBUD: TDateTimeField;
    TaEX_DATECUMBUDGET: TDateTimeField;
    EX_ETATCPTA: THDBValComboBox;
    TEX_ETATCPTA: THLabel;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private    { Déclarations privées }
    Qui : Byte ; { 0 : Menu , 1 : depuis assistant création société }
    NombreExo : Byte ;
    Function  EnregOK : boolean ; Override ;
    Procedure NewEnreg ; Override ;
    Procedure ChargeEnreg ; Override ;
    Function OnDelete : Boolean ; Override ;
    function  EcrExist(Exercice : String3 ; Table : String) : Boolean ;
    function  CtrlDebFinExo : Boolean ;
    function  CtrlDureeMinMaxExo : Boolean ;
    function  CtrlModifDateFin : Boolean ;
    procedure ActiveNewEnreg ;
    Function  CtrlDateFin : Boolean ;
    Function  CtrlDateDeb : Boolean ;
  public    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Procedure FicheExerciceSISCO(Quel : String ; Comment : TActionFiche ; Origine: Byte) ;
var FExercice: TFExoSISCO;
    PP : THPanel ;
begin
if Comment<>taConsult then
   BEGIN
   if Blocage(['nrCloture','nrBatch','nrSaisieCreat','nrSaisieModif','nrPointage','nrLettrage'],True,'nrCloture') then Exit ;
   END ;
FExercice:=TFExoSISCO.Create(Application) ;
FExercice.InitFL('EX','PRT_EXO',Quel,'',Comment,TRUE,FExercice.TaEX_EXERCICE,
                 FExercice.TaEX_LIBELLE,FExercice.TaEX_EXERCICE,['ttExercice']) ;
FExercice.Qui:=Origine ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FExercice.ShowModal ;
    finally
     FExercice.Free ;
     if Comment<>taConsult then Bloqueur('nrCloture',False) ;
    end ;
   END else
   BEGIN
   InitInside(FExercice,PP) ;
   FExercice.Show ;
   END ;
Screen.Cursor:=SyncrDefault ;
end ;

{============================================================================}
Function TFExoSISCO.EcrExist(Exercice : String3 ; Table : String) : Boolean ;
var Pref : String ;
BEGIN
Result:=False ;
if (Exercice='') then Exit ;
Pref:='E_' ;
if Table='BUDECR' then Pref:='BE_' ;
Result:=ExisteSQL('Select EX_EXERCICE From EXERCICE Where '+
                  'Exists(Select '+Pref+'EXERCICE From '+Table+' Where '+Pref+'EXERCICE="'+Exercice+'") '+
                  'And EX_EXERCICE="'+Exercice+'"') ;
END ;


{============================================================================}
Function CombienExo : Integer;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT COUNT(EX_EXERCICE) FROM EXERCICE',TRUE) ; Result:=Q.Fields[0].AsInteger ; Ferme(Q) ;
END ;


Function TFExoSISCO.EnregOK : boolean ;
Var Reponse : Integer ;
BEGIN
result:=Inherited EnregOK  ; if Not Result then Exit ;
Modifier:=True ;
if ((result) and (Ta.state in [dsEdit,dsInsert]) and (EX_DATEFIN.Enabled)) then
   BEGIN
   Result:=False ;
   if Not CtrlDateDeb then BEGIN HM2.Execute(5,'','') ; EX_DATEDEBUT.SetFocus ; Exit ; END ;
   if Not CtrlDateFin then BEGIN HM2.Execute(3,'','') ; EX_DATEFIN.SetFocus ; Exit ; END ;
   if Not CtrlDebFinExo then BEGIN HM2.Execute(0,'','') ; EX_DATEFIN.SetFocus ; Exit ; END ;
   if Not CtrlDureeMinMaxExo then BEGIN HM2.Execute(1,'','') ; EX_DATEFIN.SetFocus ; Exit ; END ;
   if Not CtrlModifDateFin then BEGIN HM2.Execute(2,'','') ; EX_DATEFIN.SetFocus ; Exit ; END ;
   END ;
Result:=True ; Modifier:=False ;
END ;

{==Pour tous les controles de validiter de date===}

Function TFExoSISCO.CtrlDebFinExo : Boolean ;
BEGIN Result:=FinDeMois(TaEX_DATEFIN.AsDateTime)>=DebutDeMois(TaEX_DATEDEBUT.AsDateTime) ; END ;

function TFExoSISCO.CtrlDureeMinMaxExo : Boolean ;
var ad,md,NbMois : Word ;
    DateExo      : TExoDate ;
BEGIN
Result:=True ;
DateExo.Deb:=DebutDeMois(TaEX_DATEDEBUT.AsDateTime) ;
DateExo.Fin:=FinDeMois(TaEX_DATEFIN.AsDateTime) ;
NombreperExo(DateExo,md,ad,NbMois) ;
if(NbMois<1)or(NbMois>23) then Result:=False ;
END ;



{============================================================================}
Function TFExoSISCO.CtrlModifDateFin : Boolean ;
BEGIN Result:=Not EcrExist(FormatFloat('000',StrToInt(Ta.FindField('EX_EXERCICE').AsString)+1),'ECRITURE') ; END ;

Function TFExoSISCO.CtrlDateDeb : Boolean ;
BEGIN Result:=(DebutDeMois(TaEX_DATEDEBUT.AsDateTime)=TaEX_DATEDEBUT.AsDateTime) ; END ;

Function TFExoSISCO.CtrlDateFin : Boolean ;
BEGIN Result:=(FinDemois(DebutDeMois(TaEX_DATEFIN.AsDateTime))=TaEX_DATEFIN.AsDateTime) ; END ;

{============================================================================}
procedure TFExoSISCO.ActiveNewEnreg ;
var NumExo    : integer ;
    DatDebut  : TDateTime ;
    DatFin    : TDateTime ;
    a,m,j     : Word ;
begin
case NombreExo of
     0 : BEGIN
         EX_EXERCICE.Text:=FormatFloat('000',001) ;
         TaEX_EXERCICE.AsString:=FormatFloat('000',001) ;
         EX_LIBELLE.Enabled:=True ;   EX_ABREGE.Enabled:=True;
         EX_DATEDEBUT.Enabled:=True ; EX_DATEFIN.Enabled:=True ;
         EX_ETATCPTA.Value:='OUV' ; TaEX_NATEXO.AsString:='' ; 
         END ;
   else  BEGIN
         TExercice.Open ; TExercice.Last ;
         DatDebut:=TExerciceEX_DATEFIN.asDateTime+1 ;
         DatFin:=TExerciceEX_DATEFIN.asDateTime;
         DecodeDate(DatFin,a,m,j) ; Inc(a) ;
         DatFin:=EncodeDate(a,m,j);
         NumExo:=StrToInt(TExercice.FindField('EX_EXERCICE').asString) ; Inc(NumExo) ;
         EX_EXERCICE.Text:=FormatFloat('000',NumExo) ;
         EX_DATEDEBUT.Enabled:=False ; EX_DATEFIN.Enabled:=True ;
         EX_DATEDEBUT.Text:=DateToStr(DatDebut) ;
         EX_DATEFIN.Text:=DateToStr(DatFin) ; TExercice.Close ;
         TaEX_EXERCICE.AsString:=FormatFloat('000',NumExo) ; TaEX_DATEDEBUT.AsDateTime:=DatDebut ;
         TaEX_DATEFIN.AsDateTime:=DatFin ;
         END ;
   End ;
EX_LIBELLE.SetFocus ;
end;

{============================================================================}
Procedure TFExoSISCO.NewEnreg ;
BEGIN
Inherited ;
taEX_ETATCPTA.Value:='NON' ;
taEX_ETATADV.Value:='NON' ;
taEX_ETATAPPRO.Value:='NON' ;  taEX_ETATPROD.Value:='NON' ;
taEX_ETATBUDGET.Value:='NON' ; TaEX_NATEXO.AsString:='' ;
ActiveNewEnreg ; 
EX_EXERCICE.Enabled:=False ;
END ;

Procedure TFExoSISCO.ChargeEnreg ;
BEGIN
Inherited ;
NombreExo:=CombienExo ;
if(NombreExo>1)And(Ta.State<>dsInsert) then
   EX_DATEFIN.Enabled:=not EcrExist(FormatFloat('000',StrToInt(EX_EXERCICE.Text)+1),'ECRITURE') ;
if NombreExo=1 then EX_DATEFIN.Enabled:=True ;
if EX_DATEFIN.Enabled then EX_DATEFIN.Color:=clWindow
                      else EX_DATEFIN.Color:=clBtnFace ;
END ;

Function TFExoSISCO.OnDelete : Boolean ;
BEGIN
If ta.State=dsInsert Then BEGIN Result:=FALSE ; Exit ; END ;
Result:=Inherited OnDelete ; If Not Result Then Exit ;
If EcrExist(TaEX_EXERCICE.AsString,'ECRITURE') Then BEGIN Result:=FALSE ; HM2.Execute(7,'','') ; END ;
END ;

procedure TFExoSISCO.FormShow(Sender: TObject);
begin
  inherited;
//if(Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult)then BinsertClick(Nil) ;
case NombreExo of
  0 :  BEGIN
       EX_DATEDEBUT.Enabled:=True ; EX_DATEDEBUT.Color:=clWindow ;
       EX_DATEFIN.Enabled:=True ; EX_DATEFIN.Color:=clWindow ;
       END;
  1 :  BEGIN
       EX_DATEDEBUT.Enabled:=not EcrExist('001','ECRITURE') ; EX_DATEFIN.Enabled:=True ;
       if EX_DATEDEBUT.Enabled then EX_DATEDEBUT.Color:=clWindow
                               else EX_DATEDEBUT.Color:=clBtnFace ;
       END ;
  else BEGIN
       EX_DATEDEBUT.Enabled:=False ; EX_DATEDEBUT.Color:=clBtnFace ;
       if EX_DATEFIN.Enabled then EX_DATEFIN.Color:=clWindow
                             else EX_DATEFIN.Color:=clBtnFace ;
       END ;
  end ;
end;

procedure TFExoSISCO.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then
   BEGIN
   if FTypeAction<>taConsult then Bloqueur('nrCloture',False) ;
   END ;
ChargemagExo(False) ;
ExoOk(TRUE) ;
  inherited;
end;

function ExoOk(AvecMess : Boolean) : Boolean ;
Var Q : TQuery ;
    OldD1,OldD2,D1,D2 : TDateTime ;
    i : Integer ;
    CodeExoFaux,CodeExo : String ;
    St : String ;
    ExoEnCours,ExoSuivant : TExoDate ;
    Etat : String ;
BEGIN
Result:=FALSE ;
Q:=OpenSQl('Select * From exercice order by EX_DATEDEBUT',TRUE) ;
OldD1:=0 ; OldD2:=0 ; i:=0 ; CodeExoFaux:='' ; St:='' ;
Fillchar(ExoEnCours,SizeOf(ExoEnCours),#0) ; Fillchar(ExoSuivant,SizeOf(ExoSuivant),#0) ;
While not Q.Eof Do
  BEGIN
  D1:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
  D2:=Q.FindField('EX_DATEFIN').AsDateTime ;
  CodeExo:=Q.FindField('EX_EXERCICE').AsString ;
  If (OldD1<>0) And (OldD2<>0) Then If D1<>OldD2+1 Then BEGIN i:=1 ; CodeExoFaux:=CodeExo; END ;
  If DebutDeMois(D1)<>D1 Then BEGIN i:=2 ; CodeExoFaux:=CodeExo ; END ;
  If FinDeMois(D2)<>D2 Then BEGIN i:=3 ; CodeExoFaux:=CodeExo ; END ;
  Etat:=Q.FindField('EX_ETATCPTA').AsString ;
  If (Etat='OUV') Or (Etat='CPR') Then
    BEGIN
    If ExoEnCours.Code='' Then BEGIN ExoEnCours.Code:=CodeExo ; ExoEnCours.Deb:=D1 ; ExoEnCours.Fin:=D2 ; END Else
      If (ExoSuivant.Code='') And (Etat='OUV') Then BEGIN ExoSuivant.Code:=CodeExo ; ExoSuivant.Deb:=D1 ; ExoSuivant.Fin:=D2 ; END Else
         BEGIN
         CodeExoFaux:=CodeExo ;
         If (Etat<>'OUV') And (ExoSuivant.Code='') Then i:=4 Else i:=5 ;
         END ;
    END Else
    BEGIN
    If ExoEnCours.Code<>'' Then BEGIN CodeExoFaux:=CodeExo ; i:=5 ; END ;
    END ;
  OldD1:=D1 ; OldD2:=D2 ;
  Q.Next ;
  END ;
Ferme(Q) ;
If i=0 Then Result:=TRUE Else If AvecMess Then
  BEGIN
  Case i Of
   1 : St:='La date de début d''exercice ne coincide pas avec la date de fin de l''exercice précédent' ;
   2 : St:='La date de début d''exercice n''est pas une date de début de mois' ;
   3 : St:='La date de fin d''exercice n''est pas une date de fin de mois' ;
   4 : St:='L''état du 2° exercice ouvert est incorrect' ;
   5 : St:='Etat de l''exercice incorrect : Il existe déjà deux exercices ouvert' ;
   End ;
  HShowMessage('0;ATTENTION ! Paramétrage de l''exercice '+CodeExoFaux+' incorrect;'+St+';E;O;O;O;','','') ;
  END ;
If Result Then
  BEGIN
  If ExoEnCours.Code<>'' Then
    BEGIN
{$IFDEF SPEC302}
    ExecuteSQL('UPDATE SOCIETE SET SO_EXOV8="'+ExoEnCours.Code+'" ') ;
{$ELSE}
    SetParamSoc('SO_EXOV8',ExoEnCours.Code) ; 
{$ENDIF}
    VH^.ExoV8.Code:=ExoEnCours.Code ;
    VH^.ExoV8.Deb:=ExoEnCours.Deb ;
    VH^.ExoV8.Fin:=ExoEnCours.Fin ;
    END ;
  END ;
END ;

end.
