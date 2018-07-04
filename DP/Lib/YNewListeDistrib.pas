unit YNewListeDistrib;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, HSysMenu, HTB97, StdCtrls, Hctrls, Grids, ExtCtrls, Utob, HmsgBox,
  HEnt1,YNewEmail,YRechercherDestinataire, AnnOutils, Mask, UtilMessages,
  YMessageDetailDestinataire;

type
  TFNewListeDistrib = class(TFVierge)
    Panel1: TPanel;
    GrilleElementListeDistrib: THGrid;
    GroupBox1: TGroupBox;
    HLabel1: THLabel;
    HLabel2: THLabel;
    CheckBoxPrive: TCheckBox;
    DescriptionListe: TEdit;
    BAjouterDest: TToolbarButton97;
    NomListe: THCritMaskEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
    procedure ModifierEmail (Sender: Tobject);
    procedure bAjouterDestClick(Sender: TObject);
    procedure NomListeKeyPress(Sender: TObject; var Key: Char);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BDeleteClick(Sender: TObject);
  private
   Nom         : String;
   Description : String;
   Utilisateur : String;

   NouvelleListe          : Boolean;
   TobElementListeDistrib : Tob;

   procedure InitialiserGrilleElementListeDistrib ();
  end;

//-----------------------------------------------
//--- Déclaration procedure / fonction globale
//-----------------------------------------------
function LancerNewListeDistrib (Nom : String; Description : String; Utilisateur : String; FicheCreation : Boolean) : Boolean;

implementation

{$R *.DFM}

//------------------------------------
//--- Nom   : LancerNewListeDistrib
//--- Objet :
//------------------------------------
function LancerNewListeDistrib (Nom : String; Description : String; Utilisateur : String; FicheCreation : Boolean) : Boolean;
var FNewListeDistrib: TFNewListeDistrib;
begin
 FNewListeDistrib:= TFNewListeDistrib.Create(Application);
 FNewListeDistrib.Nom:=Nom;
 FNewListeDistrib.Description:=Description;
 FNewListeDistrib.Utilisateur:=Utilisateur;
 FNewListeDistrib.NouvelleListe:=FicheCreation;

 try
  Result:=(FNewListeDistrib.ShowModal = mrOk);
 finally
  FNewListeDistrib.Free;
 end;
end;

//---------------------------------------------
//--- Nom   : FormShow
//--- Objet :
//---------------------------------------------
procedure TFNewListeDistrib.FormShow(Sender: TObject);
begin
 inherited;

 NomListe.Text:=Nom;
 DescriptionListe.Text:=Description;

 if (Nom<>'') then NomListe.Enabled:=False; // else CheckBoxPrive.checked:=True;
 if Utilisateur<>'' then CheckBoxPrive.checked:=True;

 if (TobElementListeDistrib=nil) then
  InitialiserGrilleElementListeDistrib ();
end;

//---------------------------------------------
//--- Nom   : FormClose
//--- Objet :
//---------------------------------------------
procedure TFNewListeDistrib.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 inherited;
 if (TobElementListeDistrib<>nil) then
  TobElementListeDistrib.free;
end;

//---------------------------------------------------
//--- Nom   : InitialiserGrilleElementListeDistrib
//--- Objet :
//---------------------------------------------------
procedure TFNewListeDistrib.InitialiserGrilleElementListeDistrib ();
var SSql    : String;
    Largeur : Integer;
begin
 GrilleElementListeDistrib.VidePile (False);
 if (TobElementListeDistrib=nil) then TobElementListeDistrib:=TOB.Create('YLISTEDISTRIBADR', Nil, -1);

 if (TobElementListeDistrib<>nil) then
  begin
   //--- Création de la requête
   SSql:='SELECT YLA_LISTEDISTRIB, YLA_EMAILADDRESS, YLA_EMAILNAME FROM YLISTEDISTRIBADR WHERE YLA_LISTEDISTRIB="'+Nom+'" ORDER BY YLA_EMAILADDRESS';
   TobElementListeDistrib.LoadDetailDbFromSQL('YLISTEDISTRIBADR',SSql);
   TobElementListeDistrib.PutGridDetail(GrilleElementListeDistrib, False, False, 'YLA_EMAILADDRESS;YLA_EMAILNAME');
  end;

 //--- Définition de la présentation de la grille
 Largeur := GrilleElementListeDistrib.Width div 2;
 GrilleElementListeDistrib.ColWidths[0] := Largeur;  // Adresse Email ou code utilisateur
 GrilleElementListeDistrib.ColWidths[1] := Largeur;  // Nom Commun
 HMTrad.ResizeGridColumns(GrilleElementListeDistrib);
end;

//---------------------------------------------------
//--- Nom   : BValiderClick
//--- Objet :
//---------------------------------------------------
procedure TFNewListeDistrib.BValiderClick(Sender: TObject);
var SUser       : String;
    Indice : Integer;
begin

 if (NomListe.Text='') then
  begin
   PgiInfo ('Vous devez saisir un nom de liste de distribution.',TitreHalley);
   NomListe.SetFocus;
   Exit;
  end;

 if (Length (NomListe.Text)<4) then
  begin
   PgiInfo ('Vous devez saisir un nom de liste de distribution supérieur à 3 caractères.',TitreHalley);
   NomListe.SetFocus;
   Exit;
  end;

 if (TobElementListeDistrib.Detail.Count=0) then
  begin
   PgiInfo ('La liste de distribution ne peut pas être validée car elle ne comporte aucun destinataire.',TitreHalley);
   Exit;
  end;

 if (ExisteSql ('SELECT 1 FROM YLISTEDISTRIB WHERE YLI_LISTEDISTRIB="'+NomListe.Text+'"') and (NouvelleListe)) then
  begin
   PgiInfo ('Cette liste de distribution existe déjà. Veuillez modifier le nom de la liste.',TitreHalley);
   Exit;
  end;

 if (CheckBoxPrive.checked) then SUser:=V_PGI.User else SUSer:='';

 //--- Enregistrement des données de la table YLISTEDISTRIB + YLISTEDISTRIBADR
 if (ExisteSql ('SELECT 1 FROM YLISTEDISTRIB WHERE YLI_LISTEDISTRIB="'+NomListe.Text+'"')) then
  begin
   //--- Suppression des données de la table YLISTEDISTRIB
   ExecuteSQL ('DELETE FROM YLISTEDISTRIBADR WHERE YLA_LISTEDISTRIB="'+Nom+'"');
   ExecuteSQL ('UPDATE YLISTEDISTRIB SET YLI_USER="'+SUser+'", YLI_LIBELLELISTE="'+DescriptionListe.Text+'" WHERE YLI_LISTEDISTRIB="'+NomListe.Text+'"');
   for Indice:=0 to TobElementListeDistrib.Detail.Count-1 do
    TobElementListeDistrib.Detail [Indice].SetAllModifie (True);
   TobElementListeDistrib.InsertOrUpDateDB;
  end
 else
  begin
   ExecuteSQl ('INSERT INTO YLISTEDISTRIB (YLI_LISTEDISTRIB,YLI_USER,YLI_LIBELLELISTE) VALUES ("'+NomListe.Text+'","'+SUser+'","'+DescriptionListe.Text+'")');
   for Indice:=0 to TobElementListeDistrib.Detail.Count-1 do
    begin
     TobElementListeDistrib.Detail[Indice].PutValue ('YLA_LISTEDISTRIB',NomListe.Text);
     TobElementListeDistrib.Detail[Indice].ChargeCle1;
    end;
   TobElementListeDistrib.InsertOrUpdateDB;
  end;

 ModalResult:=MrOk;
 inherited;
end;

//---------------------------------------------------
//--- Nom   : BInsertClick
//--- Objet :
//---------------------------------------------------
procedure TFNewListeDistrib.BinsertClick(Sender: TObject);
var SEmail, SNom : String;
    UneTobEnreg  : Tob;
begin
 inherited;
 SEmail:=''; SNom:='';
 if (LancerNewEmail (SEmail,SNom)) then
  begin
   if (SEmail<>'') then
    begin
     UneTobEnreg:=TOB.Create('YLISTEDISTRIBADR', TobElementListeDistrib, -1);
     UneTobEnreg.InitValeurs;
     UneTobEnreg.PutValue ('YLA_LISTEDISTRIB',NomListe.Text);
     UneTobEnreg.PutValue ('YLA_EMAILADDRESS',SEmail);
     UneTobEnreg.PutValue ('YLA_EMAILNAME',SNom);
     GrilleElementListeDistrib.VidePile (False);
     TobElementListeDistrib.PutGridDetail(GrilleElementListeDistrib, False, False, 'YLA_EMAILADDRESS;YLA_EMAILNAME')
    end;
  end;
end;

//---------------------------------------------
//--- Nom   : ModifierListeDistrib
//--- Objet :
//---------------------------------------------
procedure TFNewListeDistrib.ModifierEmail (Sender: Tobject);
var UneTob, UneTobEnreg : Tob;
    SEmail : String;
    SNom   : String;
begin
 inherited;
 UneTob:=Tob(GrilleElementListeDistrib.Objects[0,GrilleElementListeDistrib.Row]);

 if (UneTob<>nil) then
  begin
   SEmail:=UneTob.GetValue ('YLA_EMAILADDRESS');
   SNom:=UneTob.GetValue ('YLA_EMAILNAME');

   if (LancerNewEmail (SEmail,SNom)) then
    begin
     if (SEmail<>'') then
      begin
       UneTob.Free;
       UneTobEnreg:=TOB.Create('YLISTEDISTRIBADR', TobElementListeDistrib, -1);
       UneTobEnreg.InitValeurs;
       UneTobEnreg.PutValue ('YLA_LISTEDISTRIB',NomListe.Text);
       UneTobEnreg.PutValue ('YLA_EMAILADDRESS',SEmail);
       UneTobEnreg.PutValue ('YLA_EMAILNAME',SNom);
       GrilleElementListeDistrib.VidePile (False);
       TobElementListeDistrib.PutGridDetail(GrilleElementListeDistrib, False, False, 'YLA_EMAILADDRESS;YLA_EMAILNAME');
      end;
    end;
  end;
end;

//---------------------------------------------------
//--- Nom   : BDefaireClick
//--- Objet :
//---------------------------------------------------
procedure TFNewListeDistrib.bAjouterDestClick(Sender: TObject);
var ListeDestinataire : String;
    UnDestinataire    : String;
    Indice            : Integer;
    UneTobDest        : Tob;
    UneTobEnreg       : Tob;
begin
 inherited;

 ListeDestinataire:='';
 for Indice:=0 to TobElementListeDistrib.Detail.Count-1 do
  ListeDestinataire:=ListeDestinataire+TobElementListeDistrib.Detail[Indice].GetValue ('YLA_EMAILADDRESS')+'|'+TobElementListeDistrib.Detail[Indice].GetValue ('YLA_EMAILNAME')+'|'+';';

 UneTobDest:=Tob.Create ('',nil,-1);
 InitialiserTobListeDestinataire (UneTobDest,ListeDestinataire);

 if (LancerDestinataire (V_PGI.NoDossier, UneTobDest, True)) then
  begin
   for Indice:=0 to UneTobDest.Detail.Count-1 do
    begin
     UnDestinataire:=UneTobDest.Detail [Indice].GetValue ('Nom');
     if (TobElementListeDistrib.FindFirst (['YLA_EMAILADDRESS'],[UnDestinataire],True)=nil) then
      begin
       UneTobEnreg:=TOB.Create('YLISTEDISTRIBADR',TobElementListeDistrib, -1);
       UneTobEnreg.InitValeurs;
       UneTobEnreg.PutValue ('YLA_LISTEDISTRIB',NomListe.Text);
       UneTobEnreg.PutValue ('YLA_EMAILADDRESS',UnDestinataire);
       UneTobEnreg.PutValue ('YLA_EMAILNAME',UneTobDest.Detail [Indice].GetValue ('Libelle'));
      end;
    end;
   GrilleElementListeDistrib.VidePile (False);
   TobElementListeDistrib.PutGridDetail(GrilleElementListeDistrib, False, False, 'YLA_EMAILADDRESS;YLA_EMAILNAME');
  end;

 UneTobDest.Free;
end;

//-----------------------------------------------------------------------------
//--- Nom   : BDeleteClick
//--- Objet :
//-----------------------------------------------------------------------------
procedure TFNewListeDistrib.BDeleteClick(Sender: TObject);
var Indice : Integer;
begin
 Inherited;
 if (TobElementListeDistrib<>nil) and (GrilleElementListeDistrib.Objects [0,1]<>Nil) then
  begin
   if GrilleElementListeDistrib.AllSelected then
    TobElementListeDistrib.ClearDetail
   else
    begin
     for Indice:=GrilleElementListeDistrib.RowCount-1 Downto 1 do
      begin
       if (GrilleElementListeDistrib.IsSelected (Indice)) then
        TobElementListeDistrib.detail[Indice-1].Free;
      end;
    end;

   GrilleElementListeDistrib.VidePile (False);
   GrilleELementListeDistrib.AllSelected:=False;
   if (TobElementListeDistrib<>nil) then TobElementListeDistrib.PutGridDetail(GrilleElementListeDistrib,False,False,'YLA_EMAILADDRESS;YLA_EMAILNAME');
  end;
end;

//------------------------------------
//--- Nom   : NomListeKeyPress
//--- Objet :
//------------------------------------
procedure TFNewListeDistrib.NomListeKeyPress(Sender: TObject; var Key: Char);
begin
 inherited;

 if (Key in [Chr (VK_RETURN),Chr (VK_ESCAPE), chr (VK_PRIOR), Chr (VK_NEXT), chr (VK_END), chr (VK_HOME), chr (VK_UP), chr (VK_DOWN), chr (VK_BACK)]) then Exit;
 if not ((Key in ['0'..'9']) or (Key in ['a'..'z']) or (Key in ['A'..'Z']) or (Key in ['_']) or (Key in ['-'])) then key:=#0;
end;

//--------------------------
//--- Nom   : FormKeyDown
//--- Objet :
//--------------------------
procedure TFNewListeDistrib.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 inherited;
 //--- Touche CTRL + A
 if (Shift = [ssCtrl]) and (Key=65) then
  GrilleElementListeDistrib.AllSelected:=not (GrilleElementListeDistrib.AllSelected);
end;


end.

