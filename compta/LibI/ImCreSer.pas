unit ImCreSer;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, HPanel, StdCtrls, Mask, Hctrls, Buttons,LookUp,
  {$IFDEF EAGLCLIENT}
  utob,
  {$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  ImPlan;

type
  TFCreationEnSerie = class(TForm)
    HPanel1: THPanel;
    HLabel1: THLabel;
    HLabel2: THLabel;
    HLabel3: THLabel;
    IMMOADUPLIQUER: THCritMaskEdit;
    NOMBRE: THNumEdit;
    HPanel2: THPanel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    lNUMEROCAS: THLabel;
    NUMEROCAS: THNumEdit;
    CODEDEPART: THNumEdit;
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure IMMOADUPLIQUERElipsisClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

procedure LanceCreationEnSerie;

implementation

{$R *.DFM}

procedure LanceCreationEnSerie;
var
  FCreationEnSerie: TFCreationEnSerie;
begin
  FCreationEnSerie:= TFCreationEnSerie.Create(Application);
  try
   FCreationEnSerie.ShowModal;
  finally
    FCreationEnSerie.Free;
  end;
end;

procedure TFCreationEnSerie.BitBtn2Click(Sender: TObject);
begin
  ModalResult := mrYes;
end;

procedure TFCreationEnSerie.BitBtn1Click(Sender: TObject);
var QOrig,QDest : TQuery;
    i,j : integer;
begin
   QOrig := OpenSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+IMMOADUPLIQUER.Text+'"',True);
   if not QOrig.Eof then
   begin
     for i := 1 to Trunc(NOMBRE.Value) do
     begin
       QDest:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+W_W+'"', FALSE);
       QDest.Insert ; InitNew(QDest) ;
       for j:=0 to QOrig.FieldCount-1 do
         QDest.Fields[j].AsVariant:=QOrig.Fields[j].AsVariant ;
       QDest.FindField('I_IMMO').AsString := Format('%.10d',[Trunc(CODEDEPART.Value)+(i-1)]);
       QDest.FindField('I_LIBELLE').AsString := 'CAS N°'+IntToStr(Trunc(NUMEROCAS.Value)+(i-1));
       QDest.FindField('I_CHANGECODE').AsString := '';
       QDest.FindField('I_IMMOLIE').AsString := '';
       QDest.FindField('I_IMMOLIEGAR').AsString := '';
       QDest.FindField('I_PLANACTIF').AsInteger := RecalculSauvePlan(QDest);
       QDest.Post;
       Ferme(QDest);
     end;
     Ferme(QOrig);
   end;
   ModalResult := mrYes;
end;

procedure TFCreationEnSerie.IMMOADUPLIQUERElipsisClick(Sender: TObject);
begin
  LookupList(TControl(Sender),'Immobilisations','IMMO','I_IMMO','I_LIBELLE','','I_IMMO', True,0)  ;
end;

end.
