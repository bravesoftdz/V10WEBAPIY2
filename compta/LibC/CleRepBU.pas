{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 04/01/2005
Modifié le ... :   /  /
Description .. : Portage en eAGL
Mots clefs ... :
*****************************************************************}
unit CleRepBU;

interface

uses
  Windows, Classes, Controls, Forms,  StdCtrls, Hctrls, Buttons, ExtCtrls,
{$IFDEF EAGLCLIENT}
  UTob,
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {JP 21/11/05 : FQ 16521 : On utilise la TOM en 2/3
   RepartBU,     // ParamRepartBUDGET}
{$ENDIF}
  RepartBU_TOM, // ParamRepartBUDGET
  HSysMenu;

Procedure GetCleRepart ( Budget : String ; ListeCol : HTStrings ; Var CleRep : String ; Var ColRef : integer ; Var Courant : boolean ) ;

type
  TFCleRepBu = class(TForm)
    HLabel1: THLabel;
    HLabel2: THLabel;
    HMTrad: THSystemMenu;
    Pied: TPanel;
    BValide: THBitBtn;
    BAbandon: THBitBtn;
    BAide: THBitBtn;
    FCleRep: THValComboBox;
    FColRef: THValComboBox;
    BZoom: THBitBtn;
    CCourant: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure BZoomClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    CleRep,BudGet : String ;
    ListeCol      : HTStrings ;
    ColRef        : integer ;
    Courant       : boolean ;
  public
  end;

implementation

{$R *.DFM}

Procedure GetCleRepart ( Budget : String ; ListeCol : HTStrings ; Var CleRep : String ; Var ColRef : integer ; Var Courant : boolean ) ;
Var X  : TFCleRepBu ;
BEGIN
CleRep:='' ; ColRef:=0 ;
X:=TFCleRepBU.Create(Application) ;
 Try
  X.Budget:=Budget ; X.ListeCol:=ListeCol ; X.Courant:=Courant ;
  if X.ShowModal=mrOk then BEGIN CleRep:=X.CleRep ; ColRef:=X.ColRef ; Courant:=X.Courant ; END ;
 Finally
  X.Free ;
 End ;
END ;


procedure TFCleRepBu.FormShow(Sender: TObject);
Var Q : TQuery ;
    i : integer ;
begin
Q:=OpenSQL('SELECT BR_REPARPERIODE, BR_LIBELLE from REPARTBU Where BR_BUDJAL="'+Budget+'"',True) ;
While Not Q.EOF do
   BEGIN
   FCleRep.Values.Add(Q.Fields[0].AsString) ;
   FCleRep.Items.Add(Q.Fields[1].AsString) ;
   Q.Next ;
   END ;
Ferme(Q) ;
for i:=1 to ListeCol.Count-1 do FColRef.Items.Add(ListeCol[i]) ;
ColRef:=0 ; CleRep:='' ;
if FCleRep.Items.Count>0 then FCleRep.ItemIndex:=0 ;
if FColRef.Items.Count>0 then FColRef.ItemIndex:=0 ;
CCourant.Checked:=Courant ; 
end;

procedure TFCleRepBu.BValideClick(Sender: TObject);
begin
CleRep:=FCleRep.Value ;
ColRef:=FColRef.ItemIndex+1 ;
Courant:=CCourant.Checked ;
end;

procedure TFCleRepBu.BZoomClick(Sender: TObject);
begin
ParamRepartBUDGET(FCleRep.Value,False);
end;

procedure TFCleRepBu.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
