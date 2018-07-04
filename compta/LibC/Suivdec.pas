unit SuivDec;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, HSysMenu, Buttons,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  hCtrls, hmsgbox, Ent1, HEnt1,
  ExtCtrls, HPanel, UiUtil, HTB97 ;

procedure CircuitDec ;

type
  TFSuivDec = class(TForm)
    Msg: THMsgBox;
    HPB: TToolWindow97;
    Dock: TDock97;
    BValider: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    Baide: TToolbarButton97;
    PFen: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    e1: TEdit;
    e2: TEdit;
    e3: TEdit;
    e4: TEdit;
    e5: TEdit;
    b1: TCheckBox;
    b2: TCheckBox;
    b3: TCheckBox;
    b4: TCheckBox;
    b5: TCheckBox;
    c1: TCheckBox;
    c2: TCheckBox;
    c3: TCheckBox;
    c4: TCheckBox;
    c5: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure eChange(Sender: TObject);
    procedure cClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BaideClick(Sender: TObject);
  private
    { Déclarations privées }
    procedure ChargeCircuit ;
    function SauveCircuit : boolean ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

procedure CircuitDec ;
var X  : TFSuivDec ;
    PP : THPanel ;
begin
X:=TFSuivDec.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     X.ShowModal ;
    finally
     X.Free ;
    end ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
end ;

procedure TFSuivDec.ChargeCircuit ;
var Q : TQuery ;
    sLib,sAbr : string ;
    n : integer ;
    C : TCheckBox ;
    E : TEdit ;
    B : TCheckBox ;
begin
//Lecture du circuit de validation
Q := OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="CID" ORDER BY CC_CODE',True) ;
n:=0 ;
while not Q.EOF do
   begin
   Inc(n) ;
   sLib := Q.FindField('CC_LIBELLE').AsString ;
   sAbr := Q.FindField('CC_ABREGE').AsString ;
   C := TCheckBox(FindComponent('c'+IntToStr(n))) ;
   E := TEdit(FindComponent('e'+IntToStr(n))) ;
   B := TCheckBox(FindComponent('b'+IntToStr(n))) ;
   C.Checked := (Trim(sLib)<>'') ;
   E.Text := sLib ;
   B.Checked := (sAbr='X') ;
   Q.Next ;
   end ;
Ferme(Q) ;
end ;

function TFSuivDec.SauveCircuit : boolean ;
var Q   : TQuery ;
    n,i : integer ;
    C,B : TCheckBox ;
    E   : TEdit ;
    f,ok: Boolean ;
begin
// Contrôle de la validité du circuit
// Vérification qu'aucun moins 2 circuit ont été paramétrés et contrôle du libellé
n:=0 ;
for i:=1 to 5 do
   begin
   C := TCheckBox(FindComponent('c'+IntToStr(i))) ;
   if C<>nil then
      begin
      if C.Checked then
         begin
         Inc(n) ;
         E:=TEdit(FindComponent('e'+IntToStr(i))) ;
         if E<>nil then
            begin
            if Trim(E.Text)='' then begin Msg.Execute(2,caption,'') ; result:=False ; exit ; end ;
            end ;
         end ;
      end ;
   end ;
if n=1 then begin Msg.Execute(0,caption,'') ; result:=False ; exit ; end ;
// Vérification des trous
f:=False ; ok:=True ;
for i:=5 downto 1 do
   begin
   C := TCheckBox(FindComponent('c'+IntToStr(i))) ;
   if C<>nil then
      begin
      if C.Checked then
         begin
         f:=True ;
         end else
         begin
         if f then ok:=False ;
         end ;
      end ;
   end ;
if not ok then begin Msg.Execute(1,caption,'') ; result:=False ; exit ; end ;
// Lecture/Ecriture du circuit de validation
SourisSablier ;
Q := OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="CID" ORDER BY CC_CODE',False) ;
n:=0 ;
while not Q.EOF do
   begin
   Inc(n) ;
   C := TCheckBox(FindComponent('c'+IntToStr(n))) ;
   E := TEdit(FindComponent('e'+IntToStr(n))) ;
   B := TCheckBox(FindComponent('b'+IntToStr(n))) ;
   Q.Edit ;
   if C.Checked then
      begin
      Q.FindField('CC_LIBELLE').AsString := E.Text ;
      if B.Checked then Q.FindField('CC_ABREGE').AsString := 'X' else Q.FindField('CC_ABREGE').AsString := '-' ;
      end else
      begin
      Q.FindField('CC_LIBELLE').AsString := '' ;
      Q.FindField('CC_ABREGE').AsString := '-' ;
      end ;
   Q.Post ;
   Q.Next ;
   end ;
Ferme(Q) ;
result:=True ;
SourisNormale ;
end ;

procedure TFSuivDec.FormShow(Sender: TObject);
begin
ChargeCircuit ;
end;

procedure TFSuivDec.BValiderClick(Sender: TObject);
begin
if SauveCircuit then
   begin
   AvertirTable('CPCIRCUITDECAIS') ;
   if Not V_PGI.OutLook then ModalResult:=mrOK ;
   end ;
end;

procedure TFSuivDec.eChange(Sender: TObject);
var C : TCheckBox ;
    s : string ;
begin
s:='c'+Copy(TControl(Sender).Name,2,1) ;
C:=TCheckBox(FindComponent(s)) ;
C.Checked:=Trim(TEdit(Sender).Text)<>'' ;
end;

procedure TFSuivDec.cClick(Sender: TObject);
var E : TEdit ;
    B : TCheckBox ;
    s : string ;
begin
s:='b'+Copy(TControl(Sender).Name,2,1) ;
B:=TCheckBox(FindComponent(s)) ;
s:='e'+Copy(TControl(Sender).Name,2,1) ;
E:=TEdit(FindComponent(s)) ;
if not TCheckBox(Sender).Checked then begin E.Text:='' ; B.Checked:=False ; end ;
end;


procedure TFSuivDec.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then Action:=caFree ;
end;

procedure TFSuivDec.BaideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
