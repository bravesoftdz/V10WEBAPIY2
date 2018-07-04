unit YListeDistrib;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, HSysMenu, HTB97, Grids, Hctrls, ExtCtrls, Utob, YNewListeDistrib,
  Hent1;

type
  TFListeDistrib = class(TFVierge)
   Panel1: TPanel;
   GrilleListeDistrib: THGrid;
   procedure FormShow(Sender: TObject);
   procedure BinsertClick(Sender: TObject);
   procedure ModifierListeDistrib(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
   TobListeDistrib : Tob;
   procedure InitialiserGrilleListeDistrib ();
  end;

//------------------------------------
//--- Déclaration procedure globale
//------------------------------------
Procedure LancerListeDistrib ();

implementation

{$R *.DFM}

//---------------------------------
//--- Nom   : LancerListeDistrib
//--- Objet :
//---------------------------------
procedure LancerListeDistrib;
var FListeDistrib: TFListeDistrib;
begin
 FListeDistrib:= TFListeDistrib.Create(Application);
 try
  FListeDistrib.ShowModal;
 finally
  FListeDistrib.Free;
 end;
end;

//---------------------------------------------
//--- Nom   : FormShow
//--- Objet :
//---------------------------------------------
procedure TFListeDistrib.FormShow(Sender: TObject);
begin
 inherited;
 if (TobListeDistrib=nil) then
  InitialiserGrilleListeDistrib ();
end;

//---------------------------------------------
//--- Nom   : FormClose
//--- Objet :
//---------------------------------------------
procedure TFListeDistrib.FormClose(Sender: TObject;var Action: TCloseAction);
begin
 inherited;
 if (TobListeDistrib<>nil) then
  TobListeDistrib.free;
end;

//---------------------------------------------
//--- Nom   : InitialiserGrilleListeDistrib
//--- Objet :
//---------------------------------------------
procedure TFListeDistrib.InitialiserGrilleListeDistrib ();
var ChSql            : String;
    Largeur, Indice  : Integer;
begin
 GrilleListeDistrib.VidePile (False);
 if (TobListeDistrib=nil) then TobListeDistrib:=TOB.Create('YLISTEDISTRIB', Nil, -1);

 if (TobListeDistrib<>nil) then
  begin
   ChSql:='SELECT YLI_LISTEDISTRIB, YLI_LIBELLELISTE, YLI_USER FROM YLISTEDISTRIB WHERE YLI_USER="'+V_PGI.User+'" OR YLI_USER="" ORDER BY YLI_LISTEDISTRIB';
   TobListeDistrib.LoadDetailFromSQL(ChSql);

   for Indice:=0 to TobListeDistrib.Detail.Count-1 do
    begin
     if (TobListeDistrib.Detail [Indice].GetValue ('YLI_USER')<>'') then
      TobListeDistrib.Detail [Indice].AddChampSupValeur ('PRIVE','#ICO#3')
     else
      TobListeDistrib.Detail [Indice].AddChampSupValeur ('PRIVE','');
    end;

   TobListeDistrib.PutGridDetail(GrilleListeDistrib, False, False, 'YLI_LISTEDISTRIB;YLI_LIBELLELISTE;PRIVE');
  end;

 //--- Définition de la présentation de la grille
 Largeur := GrilleListeDistrib.Width div 8;
 GrilleListeDistrib.ColWidths[0] := 3*Largeur;  // Nom de la liste
 GrilleListeDistrib.ColWidths[1] := 4*Largeur;  // Description
 GrilleListeDistrib.ColWidths[2] := 2*Largeur;  // Privé
 GrilleListeDistrib.ColAligns[2] := TaCenter;
 HMTrad.ResizeGridColumns(GrilleListeDistrib);
end;

//---------------------------------------------
//--- Nom   : BInsertClick
//--- Objet :
//---------------------------------------------
procedure TFListeDistrib.BinsertClick(Sender: TObject);
begin
 inherited;
 if (LancerNewListeDistrib ('','',V_PGI.User,True)) then
  InitialiserGrilleListeDistrib ();
end;

//---------------------------------------------
//--- Nom   : ModifierListeDistrib
//--- Objet :
//---------------------------------------------
procedure TFListeDistrib.ModifierListeDistrib (Sender: Tobject);
var UneTob : Tob;
begin
 inherited;
 UneTob:=Tob(GrilleListeDistrib.Objects[0,GrilleListeDistrib.Row]);

 if (LancerNewListeDistrib (UneTob.GetValue ('YLI_LISTEDISTRIB'),UneTob.GetValue ('YLI_LIBELLELISTE'),UneTob.GetValue ('YLI_USER'),False)) then
  InitialiserGrilleListeDistrib ();
end;


end.
