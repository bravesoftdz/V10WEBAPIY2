unit Exercice;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB, StdCtrls, Hctrls, Mask, DBCtrls, hmsgbox,
  Buttons, ExtCtrls, Grids, DBGrids, HDB, Ent1, HEnt1, HCompte, HSysMenu,
  Hqry, HRegCpte, HTB97, UiUtil, HPanel ;

Procedure FicheExercice(Quel : String ; Comment : TActionFiche ; Origine: Byte) ;

type
  TFExercice = class(TFFicheListe)
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
    GrpBoxEtatExo : TGroupBox;
    TEX_ETATCPTA  : THLabel;
    TEX_ETATBUDGET: THLabel;
    TEX_ETATADV   : THLabel;
    TEX_ETATAPPRO : THLabel;
    TEX_ETATPROD  : THLabel;
    EX_ETATCPTA   : THDBValComboBox;
    EX_ETATBUDGET : THDBValComboBox;
    EX_ETATADV    : THDBValComboBox;
    EX_ETATAPPRO  : THDBValComboBox;
    EX_ETATPROD   : THDBValComboBox;
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
    EEX_ETATCPTA: TEdit;
    EEX_ETATBUDGET: TEdit;
    EEX_ETATADV: TEdit;
    EEX_ETATAPPRO: TEdit;
    EEX_ETATPROD: TEdit;
    Confirm: THMsgBox;
    TaEX_BUDJAL: TStringField;
    EX_BUDJAL: THDBCpteEdit;
    TEX_BUDJAL: TLabel;
    TaEX_DATECUMRUB: TDateTimeField;
    TaEX_NATEXO: TStringField;
    TaEX_DATECUMBUD: TDateTimeField;
    TaEX_DATECUMBUDGET: TDateTimeField;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BDeleteClick(Sender: TObject);
  private    { Déclarations privées }
    Qui : Byte ; { 0 : Menu , 1 : depuis assistant création société }
    NombreExo : Byte ;
    function  EcrExist(Exercice : String3 ; Table : String) : Boolean ;
    function  AutoriseNouvelExo : Boolean ;
    procedure TestEtatExo(Sender : TObject);
    function  CtrlDebFinExo : Boolean ;
    function  CtrlDureeMinMaxExo : Boolean ;
    function  CtrlModifDateFin : Boolean ;
    procedure ActiveNewEnreg ;
    Function  CtrlDateFin : Boolean ;
    Function  CtrlDateDeb : Boolean ;
    Function  CtrlEuroDeb : Boolean ;
    Function  AutoriseNonOuvert : Boolean ;
    Procedure AffecteEdit ;
    Procedure AutoriseBudget ;
    function ExoSuivantExist(DateFin: TDateTime): Boolean;
    function ControleDateExoOk(Exo: string; DateFin: TDateTime): boolean;
  public    { Déclarations publiques }
    Function  EnregOK : boolean ; Override ;
    Procedure NewEnreg ; Override ;
    Procedure ChargeEnreg ; Override ;
  end;

implementation

{$R *.DFM}

Uses UtilPgi ;

Procedure FicheExercice(Quel : String ; Comment : TActionFiche ; Origine: Byte) ;
var FExercice: TFExercice;
    PP : THPanel ;
begin
if Comment<>taConsult then
   BEGIN
   if _Blocage(['nrCloture','nrBatch','nrSaisieCreat','nrSaisieModif','nrPointage','nrLettrage','nrRelance','nrEnca',',nrDeca'],True,'nrCloture') then Exit ;
   END ;
FExercice:=TFExercice.Create(Application) ;
FExercice.InitFL('EX','PRT_EXO',Quel,'',Comment,TRUE,FExercice.TaEX_EXERCICE,
                 FExercice.TaEX_LIBELLE,FExercice.TaEX_EXERCICE,['ttExercice','TTEXERCICEBUDGET','TTEXOSAUFPRECEDENT']) ;
FExercice.Qui:=Origine ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FExercice.ShowModal ;
    finally
     FExercice.Free ;
     if Comment<>taConsult then _Bloqueur('nrCloture',False) ;
    end ;
   END else
   BEGIN
   InitInside(FExercice,PP) ;
   FExercice.Show ;
   END ;
Screen.Cursor:=SyncrDefault ;
end ;

{============================================================================}
Function TFExercice.EcrExist(Exercice : String3 ; Table : String) : Boolean ;
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

{============================================================================}
Function TFExercice.AutoriseNouvelExo : Boolean ;

Var Q : TQuery ;
    St : String ;
BEGIN
Result:=False ;
if VH^.Suivant.Code='' then Exit ; St:='' ;
Q:=OpenSql('Select EX_ETATCPTA From EXERCICE Where EX_EXERCICE="'+VH^.Suivant.Code+'"',True) ;
If Not Q.Eof Then St:=Q.Fields[0].AsString ; Ferme(Q) ;
if St<>'OUV' then Exit ;
Result:=True ;
{
Nb:=CombienExo ;
Q:=OpenSQL('SELECT EX_ETATCPTA,EX_ETATBUDGET,EX_ETATADV,EX_ETATAPPRO,EX_ETATPROD FROM EXERCICE',TRUE) ;
if Nb>1 then
   BEGIN
   Q.Last ; Q.Prior ;

   if ((Q.Fields[0].AsString<>'CDE')Or(Q.Fields[1].AsString<>'CDE')Or
       (Q.Fields[2].AsString<>'CDE')Or(Q.Fields[3].AsString<>'CDE')Or
       (Q.Fields[4].AsString<>'CDE'))then Result:=False ;
   END ;
Ferme(Q) ;
}
END ;

{============================================================================}

procedure TFExercice.TestEtatExo(Sender : TObject);
BEGIN
if((Ta.FindField(THDBValComboBox(Sender).Name).AsString='NON')Or
   (Ta.FindField(THDBValComboBox(Sender).Name).AsString='OUV')Or
   (Ta.FindField(THDBValComboBox(Sender).Name).AsString='CPR')Or
   (Ta.FindField(THDBValComboBox(Sender).Name).AsString=''))
   then THDBValComboBox(Sender).Enabled:=True
   else THDBValComboBox(Sender).Enabled:=False ;
END ;

Function TFExercice.EnregOK : boolean ;
Var Reponse : Integer ;
BEGIN
// CA - 24/04/2002 - Contrôle dates exercice OK pour enregistrement
NextControl(Self) ;
if not ControleDateExoOk (TA.FindField('EX_EXERCICE').AsString,TA.FindField('EX_DATEFIN').AsDateTime) then
begin
  PGIBox('Modification impossible : des écritures existent après la date de fin d''exercice.',Self.Caption);
  Result := False;
  exit;
end;
// CA - 24/04/2002
result:=Inherited EnregOK  ; if Not Result then Exit ;
If Qui=1 Then
   BEGIN
   HM2.Mess[6]:=Confirm.Mess[0]+Confirm.Mess[1]+chr(13)+chr(10)+chr(13)+chr(10)+Confirm.Mess[2]+DateToStr(TaEX_DATEDEBUT.AsDateTime)+chr(13)+chr(10)+
                         Confirm.Mess[3]+DateToStr(TaEX_DATEFIN.AsDateTime)+chr(13)+chr(10)+chr(13)+chr(10)+
                         Confirm.Mess[4]+chr(13)+chr(10)+chr(13)+chr(10) ;
   Reponse:=HM2.Execute(6,'','') ;
   Case Reponse of
     mrYes :  ;
     mrNo,mrCancel  : BEGIN Result:=FALSE ; Exit ; END ;
    end ;
   END ;
Modifier:=True ;
if ((result) and (Ta.state in [dsEdit,dsInsert]) and (EX_DATEFIN.Enabled)) then
   BEGIN
   Result:=False ;
   if Not CtrlDateDeb then BEGIN HM2.Execute(5,'','') ; EX_DATEDEBUT.SetFocus ; Exit ; END ;
   if Not CtrlEuroDeb then BEGIN HM2.Execute(7,'','') ; EX_DATEDEBUT.SetFocus ; Exit ; END ;
   if Not CtrlDateFin then BEGIN HM2.Execute(3,'','') ; EX_DATEFIN.SetFocus ; Exit ; END ;
   if Not CtrlDebFinExo then BEGIN HM2.Execute(0,'','') ; EX_DATEFIN.SetFocus ; Exit ; END ;
   if Not CtrlDureeMinMaxExo then BEGIN HM2.Execute(1,'','') ; EX_DATEFIN.SetFocus ; Exit ; END ;
   if Not CtrlModifDateFin then BEGIN HM2.Execute(2,'','') ; EX_DATEFIN.SetFocus ; Exit ; END ;
   if Not AutoriseNonOuvert then Exit ;
//   if (EX_ETATBUDGET.Value<>'NON') then TaEX_NATEXO.AsString:='BUD' ;
   END ;
Result:=True ; Modifier:=False ;
END ;

{==Pour tous les controles de validiter de date===}

Function TFExercice.CtrlDebFinExo : Boolean ;
BEGIN Result:=FinDeMois(TaEX_DATEFIN.AsDateTime)>=DebutDeMois(TaEX_DATEDEBUT.AsDateTime) ; END ;

function TFExercice.CtrlDureeMinMaxExo : Boolean ;
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
Function TFExercice.CtrlModifDateFin : Boolean ;
BEGIN Result:=Not EcrExist(FormatFloat('000',StrToInt(Ta.FindField('EX_EXERCICE').AsString)+1),'ECRITURE') ; END ;

Function TFExercice.CtrlDateDeb : Boolean ;
BEGIN
Result:=(DebutDeMois(TaEX_DATEDEBUT.AsDateTime)=TaEX_DATEDEBUT.AsDateTime) ;
END ;

Function TFExercice.CtrlEuroDeb : Boolean ;
BEGIN
Result:=True ; 
if ((Ta.state in [dsInsert]) and (VH^.TenueEuro) and (NombreExo<=0)) then
   Result:=(TaEX_DATEDEBUT.AsDateTime>=V_PGI.DateDebutEuro) ;
END ;

Function TFExercice.CtrlDateFin : Boolean ;
BEGIN Result:=(FinDemois(DebutDeMois(TaEX_DATEFIN.AsDateTime))=TaEX_DATEFIN.AsDateTime) ; END ;

Function LastExo(Var Exo : tExoDate) : String ;
Var Q : tQuery ;
BEGIN
Fillchar(Exo,SizeOf(Exo),#0) ;
Q:=OpenSQL('SELECT * FROM EXERCICE ORDER BY EX_DATEDEBUT',TRUE) ;
If Not Q.Eof Then
  BEGIN
  Q.Last ;
  Exo.Code:=Q.FindField('EX_EXERCICE').AsString ;
  Exo.Deb:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
  Exo.Fin:=Q.FindField('EX_DATEFIN').AsDateTime ;
  END ;
Ferme(Q) ;
END ;

{============================================================================}
procedure TFExercice.ActiveNewEnreg ;
var NumExo    : integer ;
    DatDebut  : TDateTime ;
    DatFin    : TDateTime ;
    a,m,j     : Word ;
    Exo       : tExoDate ;
begin
BInsert.Enabled:=False ;
case NombreExo of
     0 : BEGIN
         EX_EXERCICE.Text:=FormatFloat('000',001) ;
         TaEX_EXERCICE.AsString:=FormatFloat('000',001) ;
         EX_LIBELLE.Enabled:=True ;   EX_ABREGE.Enabled:=True;
         EX_DATEDEBUT.Enabled:=True ; EX_DATEFIN.Enabled:=True ;
         If Qui=1 Then BEGIN EX_ETATCPTA.Value:='OUV' ; TaEX_NATEXO.AsString:='' ; END ;
         END ;
   else  BEGIN
         (*
         TExercice.Open ; TExercice.Last ;
         DatDebut:=TExerciceEX_DATEFIN.asDateTime+1 ;
         DatFin:=TExerciceEX_DATEFIN.asDateTime;
         NumExo:=StrToInt(TExercice.FindField('EX_EXERCICE').asString) ;
         TExercice.Close ;
         *)
         LastExo(Exo) ;
         If Exo.Code<>'' Then
           BEGIN
           DatDebut:=Exo.Fin+1 ; DatFin:=Exo.Fin ;
           DecodeDate(DatFin,a,m,j) ; Inc(a) ;
           DatFin:=EncodeDate(a,m,j);
           NumExo:=StrToInt(Exo.Code) ; Inc(NumExo) ;
           EX_EXERCICE.Text:=FormatFloat('000',NumExo) ;
           EX_DATEDEBUT.Enabled:=False ; EX_DATEFIN.Enabled:=True ;
           EX_DATEDEBUT.Text:=DateToStr(DatDebut) ; EX_DATEFIN.Text:=DateToStr(DatFin) ;
           TaEX_EXERCICE.AsString:=FormatFloat('000',NumExo) ; TaEX_DATEDEBUT.AsDateTime:=DatDebut ;
           TaEX_DATEFIN.AsDateTime:=DatFin ;
           END ;
         END ;
   End ;
EX_LIBELLE.SetFocus ;
end;

{============================================================================}
Procedure TFExercice.NewEnreg ;
BEGIN
Inherited ;
taEX_ETATCPTA.Value:='NON' ;
taEX_ETATADV.Value:='NON' ;
taEX_ETATAPPRO.Value:='NON' ;  taEX_ETATPROD.Value:='NON' ;
taEX_ETATBUDGET.Value:='OUV' ; TaEX_NATEXO.AsString:='BUD' ;
ActiveNewEnreg ; AffecteEdit ;
EX_EXERCICE.Enabled:=False ;
END ;

Procedure TFExercice.ChargeEnreg ;
BEGIN
Inherited ;
NombreExo:=CombienExo ;
//TestEtatExo(EX_ETATCPTA) ;
// GP 12/03/97 :
//EX_ETATCPTA.Enabled:=TRUE ;
If EX_ETATCPTA.Value<>'NON' Then EX_ETATCPTA.Enabled:=FALSE ;

TestEtatExo(EX_ETATBUDGET) ;
//TestEtatExo(EX_ETATADV) ; TestEtatExo(EX_ETATAPPRO) ;
//TestEtatExo(EX_ETATPROD) ;
if(NombreExo>1)And(Ta.State<>dsInsert) then
   EX_DATEFIN.Enabled:=(not EcrExist(FormatFloat('000',StrToInt(EX_EXERCICE.Text)+1),'ECRITURE'))
// CA - 23/04/2002 - Pas de modification possible de la date de fin si
// un exercice suivant a déjà été créé.
    and (not ExoSuivantExist(Ta.FindField('EX_DATEFIN').AsDateTime)) ;
if NombreExo=1 then EX_DATEFIN.Enabled:=True ;
if EX_DATEFIN.Enabled then EX_DATEFIN.Color:=clWindow
                      else EX_DATEFIN.Color:=clBtnFace ;
BInsert.Enabled:=AutoriseNouvelExo ;
If (Qui=1) And (NombreExo>0) Then BInsert.Enabled:=FALSE ;
//Lhe,GG le 21/11/96
BDelete.Enabled:=(Ta.FindField('EX_ETATCPTA').AsString='NON') ;
AutoriseBudget ; AffecteEdit ;
END ;

Procedure TFExercice.AutoriseBudget ;
var Ok : boolean ;
BEGIN
Exit ;
Ok:=(TaEX_ETATCPTA.Value<>'CDE') ;
EEX_ETATBUDGET.Visible:=Not Ok ; EX_ETATBUDGET.Visible:=Ok ;
END ;

procedure TFExercice.FormShow(Sender: TObject);
begin
  inherited;
if (Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult)then
   BEGIN
   if ta.State=dsInsert then NewEnreg else BinsertClick(Nil) ;
   END ;
Case Qui Of
  1 : BEGIN (* Assistant société *)
      EEX_EtatCpta.Visible:=FALSE ; EX_EtatCpta.Visible:=TRUE ;
      EEX_ETATBUDGET.Enabled:=FALSE ;
      EEX_ETATADV.Enabled:=FALSE ;
      EEX_ETATAPPRO.Enabled:=FALSE ;
      EEX_ETATPROD.Enabled:=FALSE ;
      END ;
  END ;
case NombreExo of
  0 :  BEGIN
       EX_DATEDEBUT.Enabled:=True ; EX_DATEDEBUT.Color:=clWindow ;
       EX_DATEFIN.Enabled:=True ; EX_DATEFIN.Color:=clWindow
       END;
  1 :  BEGIN
       EX_DATEDEBUT.Enabled:=not EcrExist('001','ECRITURE') ; EX_DATEFIN.Enabled:=True ;
       if EX_DATEDEBUT.Enabled then EX_DATEDEBUT.Color:=clWindow
                               else EX_DATEDEBUT.Color:=clBtnFace ;
       If (Qui=1) And (Not Ta.Eof) And (ta.State=dsInsert) Then Bouge(nbCancel) ;
       END ;
  else BEGIN
       EX_DATEDEBUT.Enabled:=False ; EX_DATEDEBUT.Color:=clBtnFace ;
//Déjà fait dans inherited par chargeenreg
//       EX_DATEFIN.Enabled:=not EcrExist(FormatFloat('000',StrToInt(EX_EXERCICE.Text)+1)) ;
       if EX_DATEFIN.Enabled then EX_DATEFIN.Color:=clWindow
                             else EX_DATEFIN.Color:=clBtnFace ;
       END ;
end ;
If Qui=1 Then BInsert.Enabled:=(NombreExo=0) ;
If (Qui=1) and (NombreExo=1) Then
   BEGIN
   EX_ETATCPTA.Value:='OUV' ; TaEX_NATEXO.AsString:='' ;
   TaEX_ETATCPTA.AsString:='OUV' ;
   END ;
//Lhe,GG le 21/11/96
//BDelete.Enabled:=(Ta.FindField('EX_ETATCPTA').AsString='NON') ;
{$IFDEF CCS3}
EEX_ETATBUDGET.Visible:=False ; TEX_ETATBUDGET.Visible:=False ; 
{$ENDIF}
end;

Function TFExercice.AutoriseNonOuvert : Boolean ;
Var EcrCpta : Boolean ;
    EcrBud : Boolean ;
BEGIN
Result:=False ;
EcrCpta:=EcrExist(TaEX_EXERCICE.AsString,'ECRITURE') ;
EcrBud:=EcrExist(TaEX_EXERCICE.AsString,'BUDECR') ;
SourisNormale ;
if EcrCpta then
  BEGIN
  if EX_ETATCPTA.Value='NON'   then BEGIN HM2.Execute(4,'','') ; Exit ; END ;
//  if EX_ETATADV.Value='NON'    then BEGIN HM2.Execute(4,'','') ; Exit ; END ;
//  if EX_ETATAPPRO.Value='NON'  then BEGIN HM2.Execute(4,'','') ; Exit ; END ;
//  if EX_ETATPROD.Value='NON'   then BEGIN HM2.Execute(4,'','') ; Exit ; END ;
  END ;
if EcrBud then
  if EX_ETATBUDGET.Value='NON' then BEGIN HM2.Execute(4,'','') ; Exit ; END ;
Result:=True ;
END ;

Procedure TFExercice.AffecteEdit ;
BEGIN
EEX_ETATCPTA.Text:=EX_ETATCPTA.Items[EX_ETATCPTA.ItemIndex] ;
EEX_ETATBUDGET.Text:=EX_ETATBUDGET.Items[EX_ETATBUDGET.ItemIndex] ;
//EEX_ETATADV.Text:=EX_ETATADV.Items[EX_ETATADV.ItemIndex] ;
//EEX_ETATAPPRO.Text:=EX_ETATAPPRO.Items[EX_ETATAPPRO.ItemIndex] ;
//EEX_ETATPROD.Text:=EX_ETATPROD.Items[EX_ETATPROD.ItemIndex] ;
END ;

procedure TFExercice.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then
   BEGIN
   if FTypeAction<>taConsult then _Bloqueur('nrCloture',False) ;
   END ;
ChargemagExo(False) ;
  inherited;
end;

{============================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 23/04/2002
Modifié le ... :   /  /
Description .. : Vérifier si un exercice suivant a déjà été créé
Mots clefs ... :
*****************************************************************}
Function TFExercice.ExoSuivantExist(DateFin : TDateTime) : Boolean ;
BEGIN
Result:=ExisteSQL('Select EX_EXERCICE From EXERCICE Where EX_DATEDEBUT="'+
    USDateTime(DateFin+1)+'"');
END ;

Function TFExercice.ControleDateExoOk (Exo : string; DateFin : TDateTime) : boolean;
begin
  Result := not ExisteSQL('Select E_EXERCICE From ECRITURE Where E_EXERCICE="'+Exo+'" '+
                  'And E_DATECOMPTABLE>"'+USDateTime(DateFin)+'"') ;
  if Result then
    Result := not ExisteSQL('Select BE_EXERCICE From BUDECR Where BE_EXERCICE="'+Exo+'" '+
                  'And BE_DATECOMPTABLE>"'+USDateTime(DateFin)+'"') ;
end;

procedure TFExercice.BDeleteClick(Sender: TObject);
begin
  // CA - 24/04/2002 - Suppression de l'exercice uniquement si pas d'écritures
  // sur l'exercice considéré.
  if not EcrExist(Ta.FindField('EX_EXERCICE').AsString,'ECRITURE') then
  begin
    if not EcrExist(Ta.FindField('EX_EXERCICE').AsString,'BUDECR') then
      inherited;
  end else
    PGIBox ('Présence d''écritures sur cet exercice.Suppression impossible',Self.Caption);
end;

end.
