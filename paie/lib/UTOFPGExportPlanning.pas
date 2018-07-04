unit UTOFPGExportPlanning;

interface
{
PT1 27/06/2006 SB V_65 Améliorations & Refonte de l'export des compteurs du planning
}

uses  Classes,HCtrls,Dialogs,sysutils,graphics,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      FE_Main,
{$ENDIF}
      AGLInit,HTB97,UTOF,Utob,HXLSPAS;   
//
Type

     TOF_PGEXPORTPLANNING = Class (TOF)
     public
        procedure OnLoad ; override ;
        procedure OnArgument(Arguments : String ) ; override ;
     private
        LeTitre,NomColonne,TitreColonne : String;
        procedure ExportClik (Sender: TObject);
     END ;

procedure PGExportPlanning ( UnObjet : Tob; Titre, NomCol, TitreCol : String );

implementation

uses pgplanningoutils;

procedure PGExportPlanning ( UnObjet : Tob; Titre, NomCol, TitreCol : String );
Begin
TheTob:=UnObjet ;
AGLLanceFiche('PAY','EXPORTPLANNING','','',Titre+';'+NomCol+'|'+TitreCol);
End;

{ TOF_PGExportPlanning }

procedure TOF_PGEXPORTPLANNING.OnArgument(Arguments: String);
var BtnExport  : TToolbarButton97;
begin
  inherited;
LeTitre   := ReadTokenSt(Arguments);
NomColonne :=ReadTokenPipe(Arguments,'|');
TitreColonne :=Arguments;
BtnExport    := TToolbarButton97 (GetControl ('BEXPORTER'));
If BtnExport <> NIL then BtnExport.OnClick := ExportClik;
end;

procedure TOF_PGEXPORTPLANNING.OnLoad;
Var i : integer;
    LaGrille : THGrid;
begin
  inherited;
LaGrille := THGrid(GetControl('GRILLE'));

LaTob.PutGridDetail(LaGrille,True,True,NomColonne,False);
If (Ecran<>nil) and (LeTitre<>'') then Ecran.Caption:= LeTitre;
i := 0;
While TitreColonne <> '' do
  Begin
  LaGrille.ColNames[i] := ReadTokenSt(NomColonne);
  LaGrille.Titres.Add(ReadTokenSt(TitreColonne));
  inc(i);
  End;
LaGrille.UpdateTitres;

For i := RendIndexCol(LaGrille,'CPACQUISN1') to LaGrille.colCount-1 do
   Begin
   LaGrille.ColAligns[i] := taRightJustify;
   LaGrille.ColWidths[i] :=  65;
   End;
   
For i :=0  to RendIndexCol(LaGrille,'CPACQUISN1') - 1 do
  LaGrille.ColWidths[i] :=  75;

LaGrille.ColWidths[RendIndexCol(LaGrille,'NOM_SAL')] := 200;

LaGrille.ColColors[RendIndexCol(LaGrille,'CPACQUISN1')] := clBlue;
LaGrille.ColColors[RendIndexCol(LaGrille,'CPPRISN1')]   := clBlue;
LaGrille.ColColors[RendIndexCol(LaGrille,'CPRESTN1')]   := clBlue;
LaGrille.ColColors[RendIndexCol(LaGrille,'CPACQUISN')]  := clBlue;
LaGrille.ColColors[RendIndexCol(LaGrille,'CPPRISN')]    := clBlue;
LaGrille.ColColors[RendIndexCol(LaGrille,'CPRESTN')]    := clBlue;
LaGrille.ColColors[RendIndexCol(LaGrille,'CPENCOURS')]  := clBlue;
LaGrille.ColColors[RendIndexCol(LaGrille,'CPSOLDE')]    := clBlue;
LaGrille.ColColors[RendIndexCol(LaGrille,'RTTACQUIS')]  := clGreen;
LaGrille.ColColors[RendIndexCol(LaGrille,'RTTPRIS')]    := clGreen;
LaGrille.ColColors[RendIndexCol(LaGrille,'RTTREST')]    := clGreen;
LaGrille.ColColors[RendIndexCol(LaGrille,'RTTENCOURS')] := clGreen;
LaGrille.ColColors[RendIndexCol(LaGrille,'RTTSOLDE')]   := clGreen;
LaGrille.ColColors[RendIndexCol(LaGrille,'ABSENCE')]    := clPurple;

end;

procedure TOF_PGEXPORTPLANNING.ExportClik(Sender: TObject);
Var Grille : ThGrid;
    SD: TSaveDialog;
begin
Grille:=ThGrid(GetControl('GRILLE'));
SD:=TSaveDialog.Create(nil);
SD.Filter:='Fichier Texte (*.txt)|*.txt|Fichier Excel (*.xls)|*.xls|Fichier Ascii (*.asc)|*.asc|Fichier Lotus (*.wks)|*.wks';
//|Fichier HTML (*.html)|*.html|Fichier XML (*.xml)|*.xml';
if SD.Execute then
  Begin
//Fichier Texte (*.txt)|*.txt|  SD.FilterIndex = 1
//Fichier Excel (*.xls)|*.xls|  SD.FilterIndex = 2
//Fichier Ascii (*.asc)|*.asc|  SD.FilterIndex = 3
//Fichier Lotus (*.wks)|*.wks|  SD.FilterIndex = 4
  if (SD.FilterIndex=1) and (Pos('.txt',SD.FileName)<1) then
     SD.FileName:=Trim(SD.FileName)+'.txt'
  else
    if (SD.FilterIndex=2) and (Pos('.xls',SD.FileName)<1) then
       SD.FileName:=Trim(SD.FileName)+'.xls'
    else
      if (SD.FilterIndex=3) and (Pos('.asc',SD.FileName)<1) then
         SD.FileName:=Trim(SD.FileName)+'.asc'
      else
        if (SD.FilterIndex=4) and (Pos('.wks',SD.FileName)<1) then
           SD.FileName:=Trim(SD.FileName)+'.wks';
  ExportGrid(Grille,nil,SD.FileName,SD.FilterIndex,True) ;
  End;
SD.Destroy;
end;


Initialization
registerclasses([TOF_PGEXPORTPLANNING]) ;
end.
