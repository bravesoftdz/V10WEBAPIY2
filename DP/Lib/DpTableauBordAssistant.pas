//---------------------------------------------------
//--- Auteur : CATALA David
//--- Objet  : Analyse d'activitée : Assitant GI
//---------------------------------------------------
unit DpTableauBordAssistant;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Hent1,
  Grids, Hctrls, Vierge, HPanel, UIUtil,
{$IFDEF EAGLCLIENT}
  UTob,
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  Hmsgbox, DpTableauBordLibrairie;

type
  TF_Assistant = class(TForm)
   GrilleAssistant: THGrid;
   procedure FormShow(Sender: TObject);
   procedure FormCreate(Sender: TObject);
   procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
  public
   CodeAffaire    : String;
   CodeTiers      : String;
   MissionEnc     : Boolean;
   procedure InitialiserLibelleAssistant (Sender : TObject);
   procedure InitialiserCelluleAssistant (Sender : TObject);
   procedure PostDrawCell(ACol,ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
  end;

//---------------------------------
//--- Déclaration des procédures
//---------------------------------
procedure Aff_AssistantMission (CodeAffaire : String; CodeTiers : String; BMission : Boolean);

implementation

{$R *.DFM}

//------------------------------------------------
//--- Nom   : Aff_AssitantMission
//--- Objet : Affichage de la fiche Assistants Mission
//------------------------------------------------
procedure Aff_AssistantMission (CodeAffaire : String; CodeTiers : String; BMission : Boolean);
Var Fiche : TF_Assistant;
    PP : THPanel;
begin
 Fiche:=TF_Assistant.Create (Application);
 Fiche.CodeAffaire:=CodeAffaire;
 Fiche.CodeTiers:=CodeTiers;
 Fiche.MissionEnc:=BMission;
 PP:=FindInsidePanel;
 if PP=nil then
  begin
   try
    Fiche.showModal;
   finally;
    Fiche.Free;
   end;
  end
 else
  begin
   InitInside (Fiche,PP);
   Fiche.Show;
  end;
end;

//-----------------------------------
//--- Nom   : FormCreate
//--- Objet : Création de la fiche
//-----------------------------------
procedure TF_Assistant.FormCreate(Sender: TObject);
begin
 GrilleAssistant.Ctl3D:=False;
 GrilleAssistant.PostDrawCell:=PostDrawCell;
end;

//------------------------------------
//--- Nom   : FormShow
//--- Objet : Affichage de la fiche
//------------------------------------
procedure TF_Assistant.FormShow(Sender: TObject);
begin
 InitialiserLibelleAssistant (Sender);
 InitialiserCelluleAssistant (Sender);
end;

//--------------------------------------------------------
//--- Nom   : InitialiserLibelleAssistant
//--- Objet : Initialise les libellés de la fiche solde
//--------------------------------------------------------
procedure TF_Assistant.InitialiserLibelleAssistant (Sender : TObject);
var Indice : Integer;
begin
 GrilleAssistant.ColCount:=4;

 for Indice:=0 to GrilleAssistant.ColCount-1 do
  GrilleAssistant.ColWidths [Indice]:=GrilleAssistant.Width div GrilleAssistant.ColCount-1;

 GrilleAssistant.Cells [0,0]:='Code Ressource';
 GrilleAssistant.Cells [1,0]:='Nom';
 GrilleAssistant.Cells [2,0]:='Prénom';
 GrilleAssistant.Cells [3,0]:='Nombre(s) heure(s)';
end;

//--------------------------------------------------------
//--- Nom   : InitialiserCelluleSolde
//--- Objet : Initialise les cellules de la fiche solde
//--------------------------------------------------------
procedure TF_Assistant.InitialiserCelluleAssistant (Sender : TObject);
var ChSql    : String;
    RSql     : Tquery;
    NumLigne : Integer;
begin
 if (CodeAffaire<>'###TOTAUX###') then
  ChSql:='SELECT ACT_RESSOURCE,SUM (ACT_QTEUNITEREF) AS SOMMEHEURE,ARS_LIBELLE,ARS_LIBELLE2 FROM ACTIVITE left join RESSOURCE on ACT_RESSOURCE=ARS_RESSOURCE WHERE ACT_AFFAIRE="'+CodeAffaire+'" AND ACT_TYPEACTIVITE="REA" GROUP BY ACT_RESSOURCE,ARS_LIBELLE,ARS_LIBELLE2'
 else
  begin
   if (MissionEnc) then
    ChSql:='SELECT ACT_RESSOURCE,SUM (ACT_QTEUNITEREF) AS SOMMEHEURE,ARS_LIBELLE,ARS_LIBELLE2 FROM ACTIVITE '+
         'Left join RESSOURCE on ACT_RESSOURCE=ARS_RESSOURCE '+
         'Left join AFFAIRE on AFF_AFFAIRE=ACT_AFFAIRE '+
         'WHERE ACT_TIERS="'+CodeTiers+'" AND AFF_ETATAFFAIRE="ENC" AND ACT_TYPEACTIVITE="REA" '+
         'GROUP BY ACT_RESSOURCE,ARS_LIBELLE,ARS_LIBELLE2'
   else
    ChSql:='SELECT ACT_RESSOURCE,SUM (ACT_QTEUNITEREF) AS SOMMEHEURE,ARS_LIBELLE,ARS_LIBELLE2 FROM ACTIVITE '+
         'Left join RESSOURCE on ACT_RESSOURCE=ARS_RESSOURCE '+
         'Left join AFFAIRE on AFF_AFFAIRE=ACT_AFFAIRE '+
         'WHERE ACT_TIERS="'+CodeTiers+'" AND ACT_TYPEACTIVITE="REA" '+
         'GROUP BY ACT_RESSOURCE,ARS_LIBELLE,ARS_LIBELLE2'
  end;

 try
  RSql:=Opensql (ChSql, TRUE, -1, '', True);
  NumLigne:=1;

  while not (RSql.eof) do
   begin
    GrilleAssistant.Cells [0,NumLigne]:=RSql.FindField ('ACT_RESSOURCE').AsString;
    GrilleAssistant.Cells [1,NumLigne]:=RSql.FindField ('ARS_LIBELLE').AsString;
    GrilleAssistant.Cells [2,NumLigne]:=RSql.FindField ('ARS_LIBELLE2').AsString;
    GrilleAssistant.Cells [3,NumLigne]:=RSql.FindField ('SOMMEHEURE').AsString;
    inc (NumLigne);
    RSql.Next;
   end;

  Ferme (RSql);
  GrilleAssistant.RowCount:=NumLigne;

 except
  PGIINFO ('Problème dans la table Ressource','Erreur');
 end;
end;

//------------------------------------------------------------
//--- Nom   : PostDrawCell
//--- Objet : Définit le contenu de la cellule avant que la
//---         cellule soit dessiné par le canvas.
//------------------------------------------------------------
procedure TF_Assistant.PostDrawCell(ACol,ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
 if (ACol=3) and (ARow>0) then
  EcrireTexte (CENTRE,ACol,ARow,GrilleAssistant)
end;

//----------------------------------
//--- Nom   : FormKeyDown
//--- Objet : Gestion des touches
//----------------------------------
procedure TF_Assistant.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 if Key=VK_ESCAPE then
  begin
   key:=0;
   ModalResult := mrCancel;
  end;
end;

end.
