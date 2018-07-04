{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 14/04/2003
Modifié le ... : 25/06/2003
Description .. : Passage en eAGL
Suite ........ : Visuaisation des fichiers IBAN FQ11835
Mots clefs ... : 
*****************************************************************}
unit CFONBFIC;

interface

uses
  SysUtils,  // IntToStr
  Classes,
  Controls,
  Forms,
  Grids,
  StdCtrls,
  ExtCtrls,
  HCtrls,
  Buttons,
  HSysMenu, HPanel;

Procedure VisuExportCFONB (StFichier : String ) ;

type
  TFCFONBFIC = class(TForm)
    Outils: TPanel;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    PEntete: TPanel;
    HTitre: TLabel;
    HMTrad: THSystemMenu;
    pnl_IBAN: THPanel;
    G2: THGrid;
    G3: THGrid;
    G1: THGrid;
    pnl_Main: THPanel;
    G: THGrid;
    HSplitter1: THSplitter;
    HSplitter2: THSplitter;
    HSplitter3: THSplitter;
    G4: THGrid;
    procedure FormShow(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    procedure ChargeGrid ;
  public
    StFichier : String ;
  end;

implementation

{$R *.DFM}

Procedure VisuExportCFONB ( StFichier : String ) ;
Var X : TFCFONBFIC ;
BEGIN
X:=TFCFONBFIC.Create(Application) ;
 Try
  X.StFichier:=StFichier ;
  X.ShowModal ;
 Finally
  X.Free ;
 End ;
END ;

procedure TFCFONBFIC.ChargeGrid ;
Var FF   : TextFile ;
    WW   : integer ;
    St,St2   : String ;
    bIban, Okok : boolean ;
BEGIN
AssignFile(FF,StFichier) ; Okok:=False ;
{$I-} Reset(FF) ; {$I+} if IoResult<>0 then Exit ;
Readln(FF,St);
if Length(St) = 320 then begin
  bIban := True;
  pnl_IBAN.Visible := True;
  pnl_Main.Visible := False;
  
   // Largeur des colonnes
  WW:=G1.Canvas.TextWidth('W') ;
  G1.ColWidths[1]:=2*WW ;
  G1.ColWidths[2]:=2*WW ;
  G1.ColWidths[3]:=6*WW ;
  G1.ColWidths[4]:=8*WW ;
  G1.ColWidths[5]:=35*WW ;
  G1.ColWidths[6]:=105*WW ;
  G1.ColWidths[7]:=14*WW ;
  G1.ColWidths[8]:=16*WW ;
  G1.ColWidths[9]:=11*WW ;
  G1.ColWidths[10]:=1*WW ;
  G1.ColWidths[11]:=34*WW ;
  G1.ColWidths[12]:=3*WW ;
  G1.ColWidths[13]:=16*WW ;
  G1.ColWidths[14]:=1*WW ;
  G1.ColWidths[15]:=34*WW ;
  G1.ColWidths[16]:=3*WW ;
  G1.ColWidths[17]:=16*WW ;
  G1.ColWidths[18]:=1*WW ;
  G1.ColWidths[19]:=1*WW ;
  G1.ColWidths[20]:=8*WW ;
  G1.ColWidths[21]:=2*WW ;

  WW:=G2.Canvas.TextWidth('W') ;
  G2.ColWidths[1]:=2*WW ;
  G2.ColWidths[2]:=2*WW ;
  G2.ColWidths[3]:=6*WW ;
  G2.ColWidths[4]:=1*WW ;
  G2.ColWidths[5]:=34*WW ;
  G2.ColWidths[6]:=35*WW ;
  G2.ColWidths[7]:=105*WW ;
  G2.ColWidths[8]:=17*WW ;
  G2.ColWidths[9]:=2*WW ;
  G2.ColWidths[10]:=16*WW ;
  G2.ColWidths[11]:=1*WW ;
  G2.ColWidths[12]:=4*WW ;
  G2.ColWidths[13]:=14*WW ;
  G2.ColWidths[14]:=1*WW ;
  G2.ColWidths[15]:=1*WW ;
  G2.ColWidths[16]:=3*WW ;
  G2.ColWidths[17]:=2*WW ;
  G2.ColWidths[18]:=1*WW ;
  G2.ColWidths[19]:=2*WW ;
  G2.ColWidths[20]:=8*WW ;
  G2.ColWidths[21]:=34*WW ;
  G2.ColWidths[22]:=3*WW ;
  G2.ColWidths[23]:=22*WW ;
  G2.ColWidths[24]:=8*WW ;
  G2.ColWidths[25]:=2*WW ;

  WW:=G3.Canvas.TextWidth('W') ;
  G3.ColWidths[1]:=2*WW ;
  G3.ColWidths[2]:=2*WW ;
  G3.ColWidths[3]:=6*WW ;
  G3.ColWidths[4]:=35*WW ;
  G3.ColWidths[5]:=105*WW ;
  G3.ColWidths[6]:=11*WW ;
  G3.ColWidths[7]:=2*WW ;
  G3.ColWidths[8]:=157*WW ;

  WW:=G4.Canvas.TextWidth('W') ;
  G4.ColWidths[1]:=2*WW ;
  G4.ColWidths[2]:=2*WW ;
  G4.ColWidths[3]:=6*WW ;
  G4.ColWidths[4]:=8*WW ;
  G4.ColWidths[5]:=140*WW ;
  G4.ColWidths[6]:=14*WW ;
  G4.ColWidths[7]:=16*WW ;
  G4.ColWidths[8]:=11*WW ;
  G4.ColWidths[9]:=1*WW ;
  G4.ColWidths[10]:=34*WW ;
  G4.ColWidths[11]:=3*WW ;
  G4.ColWidths[12]:=16*WW ;
  G4.ColWidths[13]:=18*WW ;
  G4.ColWidths[14]:=49*WW ;
  end
else begin
  bIban := False;
  G.ColCount:=5 ;
  WW:=G.Canvas.TextWidth('W') ;
  G.ColWidths[1]:=18*WW ;
  G.ColWidths[2]:=59*WW ;
  G.ColWidths[3]:=23*WW ;
  G.ColWidths[4]:=57*WW ;
end;
{$I-} Reset(FF) ; {$I+} if IoResult<>0 then Exit ;

While Not EOF(FF) do
   BEGIN
   Readln(FF,St);

   if bIban then begin
     St2 := Copy(St,0,2);
     // En-tête
     if (St2 = '03') then begin
       G1.Cells[1,G1.RowCount-1]:=Copy(St,1,2);
       G1.Cells[2,G1.RowCount-1]:=Copy(St,3,2);
       G1.Cells[3,G1.RowCount-1]:=Copy(St,5,6);
       G1.Cells[4,G1.RowCount-1]:=Copy(St,11,8);
       G1.Cells[5,G1.RowCount-1]:=Copy(St,19,35);
       G1.Cells[6,G1.RowCount-1]:=Copy(St,54,105);
       G1.Cells[7,G1.RowCount-1]:=Copy(St,159,14);
       G1.Cells[8,G1.RowCount-1]:=Copy(St,173,16);
       G1.Cells[9,G1.RowCount-1]:=Copy(St,189,11);
       G1.Cells[10,G1.RowCount-1]:=Copy(St,200,1);
       G1.Cells[11,G1.RowCount-1]:=Copy(St,201,34);
       G1.Cells[12,G1.RowCount-1]:=Copy(St,235,3);
       G1.Cells[13,G1.RowCount-1]:=Copy(St,238,16);
       G1.Cells[14,G1.RowCount-1]:=Copy(St,254,1);
       G1.Cells[15,G1.RowCount-1]:=Copy(St,255,34);
       G1.Cells[16,G1.RowCount-1]:=Copy(St,289,3);
       G1.Cells[17,G1.RowCount-1]:=Copy(St,292,16);
       G1.Cells[18,G1.RowCount-1]:=Copy(St,308,1);
       G1.Cells[19,G1.RowCount-1]:=Copy(St,309,1);
       G1.Cells[20,G1.RowCount-1]:=Copy(St,310,8);
       G1.Cells[21,G1.RowCount-1]:=Copy(St,318,2);
       G1.RowCount:=G1.RowCount+1;
       end
     // Détail
     else if (St2 = '04') then begin
       G2.Cells[1,G2.RowCount-1]:=Copy(St,1,2);
       G2.Cells[2,G2.RowCount-1]:=Copy(St,3,2);
       G2.Cells[3,G2.RowCount-1]:=Copy(St,5,6);
       G2.Cells[4,G2.RowCount-1]:=Copy(St,11,1);
       G2.Cells[5,G2.RowCount-1]:=Copy(St,12,34);
       G2.Cells[6,G2.RowCount-1]:=Copy(St,46,35);
       G2.Cells[7,G2.RowCount-1]:=Copy(St,81,105);
       G2.Cells[8,G2.RowCount-1]:=Copy(St,186,17);
       G2.Cells[9,G2.RowCount-1]:=Copy(St,203,2);
       G2.Cells[10,G2.RowCount-1]:=Copy(St,205,16);
       G2.Cells[11,G2.RowCount-1]:=Copy(St,221,1);
       G2.Cells[12,G2.RowCount-1]:=Copy(St,222,4);
       G2.Cells[13,G2.RowCount-1]:=Copy(St,226,14);
       G2.Cells[14,G2.RowCount-1]:=Copy(St,240,1);
       G2.Cells[15,G2.RowCount-1]:=Copy(St,241,1);
       G2.Cells[16,G2.RowCount-1]:=Copy(St,242,3);
       G2.Cells[17,G2.RowCount-1]:=Copy(St,245,2);
       G2.Cells[18,G2.RowCount-1]:=Copy(St,247,1);
       G2.Cells[19,G2.RowCount-1]:=Copy(St,248,2);
       G2.Cells[20,G2.RowCount-1]:=Copy(St,250,8);
       G2.Cells[21,G2.RowCount-1]:=Copy(St,251,34);
       G2.Cells[22,G2.RowCount-1]:=Copy(St,285,3);
       G2.Cells[23,G2.RowCount-1]:=Copy(St,288,22);
       G2.Cells[24,G2.RowCount-1]:=Copy(St,310,8);
       G2.Cells[25,G2.RowCount-1]:=Copy(St,318,2);
       G2.RowCount:=G2.RowCount+1;
       end
     // Banque bénéficiaire
     else if (St2 = '05') then begin
       G3.Cells[1,G3.RowCount-1]:=Copy(St,1,2);
       G3.Cells[2,G3.RowCount-1]:=Copy(St,3,2);
       G3.Cells[3,G3.RowCount-1]:=Copy(St,5,6);
       G3.Cells[4,G3.RowCount-1]:=Copy(St,11,35);
       G3.Cells[5,G3.RowCount-1]:=Copy(St,46,105);
       G3.Cells[6,G3.RowCount-1]:=Copy(St,151,11);
       G3.Cells[7,G3.RowCount-1]:=Copy(St,162,2);
       G3.Cells[8,G3.RowCount-1]:=Copy(St,164,157);
       G3.RowCount:=G3.RowCount+1;       
       end
     // Total
     else if (St2 = '08') then begin
       G4.Cells[1,G4.RowCount-1]:=Copy(St,1,2);
       G4.Cells[2,G4.RowCount-1]:=Copy(St,3,2);
       G4.Cells[3,G4.RowCount-1]:=Copy(St,5,6);
       G4.Cells[4,G4.RowCount-1]:=Copy(St,11,8);
       G4.Cells[5,G4.RowCount-1]:=Copy(St,19,140);
       G4.Cells[6,G4.RowCount-1]:=Copy(St,159,14);
       G4.Cells[7,G4.RowCount-1]:=Copy(St,173,16);
       G4.Cells[8,G4.RowCount-1]:=Copy(St,189,11);
       G4.Cells[9,G4.RowCount-1]:=Copy(St,200,1);
       G4.Cells[10,G4.RowCount-1]:=Copy(St,201,34);
       G4.Cells[11,G4.RowCount-1]:=Copy(St,235,3);
       G4.Cells[12,G4.RowCount-1]:=Copy(St,238,16);
       G4.Cells[13,G4.RowCount-1]:=Copy(St,254,18);
       G4.Cells[14,G4.RowCount-1]:=Copy(St,272,49);
       G4.RowCount:=G4.RowCount+1;
     end;
     end
   else begin
     G.Cells[1,G.RowCount-1]:=Copy(St,1,18) ;
     G.Cells[2,G.RowCount-1]:=Copy(St,19,59) ;
     G.Cells[3,G.RowCount-1]:=Copy(St,79,23) ;
     G.Cells[4,G.RowCount-1]:=Copy(St,103,Length(St)-103); // 57);
     G.Cells[0,G.RowCount-1]:=IntToStr(G.RowCount-1) ;
     G.RowCount:=G.RowCount+1 ;
     Okok:=True ;
   end;
end;

CloseFile(FF) ;
if Okok then G.RowCount:=G.RowCount-1 ;
end;

procedure TFCFONBFIC.FormShow(Sender: TObject);
begin
HTitre.Caption:=HTitre.Caption+StFichier ;
ChargeGrid ;
end;

procedure TFCFONBFIC.BFermeClick(Sender: TObject);
begin
Close ;
end;

procedure TFCFONBFIC.BValiderClick(Sender: TObject);
begin
Close ;
end;

procedure TFCFONBFIC.FormClose(Sender: TObject; var Action: TCloseAction);
begin
G.VidePile(False) ;
end;

end.
