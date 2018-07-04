{***********UNITE*************************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... : 10/06/2003
Description .. : Consultation de l'enveloppe
Mots clefs ... : TOX;ENVELOPPE
*****************************************************************}
unit GCToxCtrlEnv;

interface

uses Windows,
     Messages,
     Classes,
     Forms,
     uToxClasses,
     uToxTools,
     uToxconst,
     StdCtrls,
     Controls,
     SysUtils, Grids, Hctrls, HEnt1;

  procedure DisplayEnveloppe ( F: STring ) ;

type
  TFicheEnveloppe = class(TForm)
    GroupBox1: TGroupBox;
    esite: TEdit;
    eLibelle: TEdit;
    eApplication: TEdit;
    eTypeSend: TEdit;
    DateMsg: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    GroupBox2: TGroupBox;
    NumChrono: TEdit;
    Label6: TLabel;
    Label7: TLabel;
    dSite: TEdit;
    Label8: TLabel;
    dLibelle: TEdit;
    GroupBox3: TGroupBox;
    Label9: TLabel;
    eTypeMsg: TEdit;
    G: THGrid;
    eEvent: TEdit;
    Label10: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCanResize(Sender: TObject; var NewWidth,
      NewHeight: Integer; var Resize: Boolean);
  private
    procedure Change;
    { Déclarations privées }
  public
    { Déclarations publiques }
    FileName: String ;
    Enveloppes: TCollectionEnveloppes ;
  end;

implementation

{$R *.DFM}

procedure DisplayEnveloppe ( F: STring ) ;
var X: TFicheEnveloppe;
begin
  X := TFicheEnveloppe.Create ( Application ) ;
  X.FileName := F ;
  X.Change ;
  X.ShowModal ;
  X.Free ;
end ;

procedure TFicheEnveloppe.FormCreate(Sender: TObject);
begin
  Enveloppes := TCollectionEnveloppes.Create(TCollectionEnveloppe) ;
end;

procedure TFicheEnveloppe.FormShow(Sender: TObject);
begin
  // Appel de la fonction d'empilage dans la liste des fiches
  AglEmpileFiche(Self) ;
end;

procedure TFicheEnveloppe.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Enveloppes.Free ;
  Action := caFree ;
end;

procedure TFicheEnveloppe.Change ;

  function VraiFaux(e,v,f: variant):variant ;
  begin
    if e then result:=v else result:=f;
  end ;

var
  E: TCollectionEnveloppe ;
  I: Integer ;
begin
  Caption := ' Contenu du fichier enveloppe : ' + ExtractFileName ( FileName ) ;

  ENveloppes.Clear ;
  Enveloppes.LoadEnveloppe(FileName) ;
  E:=Enveloppes.Items[0] ;

  esite.Text := E.eSite ;
  eLibelle.Text:=E.eLibelle;
  eApplication.Text:=E.eApplication ;
  eEvent.Text := E.CodeEvent ;
  eTypeSend.Text:=TypeToxSendToString(E.eTypeSend);
  DateMsg.Text:=FormatDateTime('dd/mm/yyyy hh:nn:ss',E.DateMsg);
  NumChrono.Text:=E.NumChrono ;
  dSite.Text:=E.dSite ;
  dLibelle.Text:=E.dLibelle ;
  eTypeMsg.Text:=TypeToxMessageToString(E.eTypeMsg);

  if Enveloppes.Items[0].LesFichiers.Count>0 then
    begin
    G.RowCount:=G.FixedRows+E.LesFichiers.Count ;
    for i:=0 to E.LesFichiers.Count-1 do
      begin
      G.Cells[0,i+G.FixedRows]:=E.LesFichiers.Items[I].FileName ;
      G.Cells[1,i+G.FixedRows]:=Format('%d',[E.LesFichiers.Items[I].Crc]) ;
      end ;
    end ;
end ;

procedure TFicheEnveloppe.FormCanResize(Sender: TObject; var NewWidth, NewHeight: Integer; var Resize: Boolean);
begin
  Resize := False ;
end;

procedure TFicheEnveloppe.FormDestroy(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche ;
end;

end.
