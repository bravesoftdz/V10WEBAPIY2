unit UAdminAgricole;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Vierge, HSysMenu, HTB97, ComCtrls,
  HDTLinks, HCtrls, hmsgbox, HEnt1, Paramsoc,  uiutil, MajTable,
  ImgList, StdCtrls, HPop97, Menus,  ExtCtrls,
  EntDP, HPanel, Utob, HImgList;

type

  TAgricoleNode = class(TObject)
   Niveau   : Integer;
   Code     : Integer;
   Libelle  : String;
   IconId   : Integer;
   IconSel  : Integer;
  end;

  TFAdminAgricole = class(TFVierge)
    TV: TTreeView;
    ImageList: THImageList;

    procedure FormShow(Sender: TObject);
    procedure AlimenterTVAgricole;
    procedure BDeleteClick(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);

   private
    function UtilisationDossier (UnCode : Integer) : Boolean;
    function DonnerNomDetail (var Libelle : String) : TModalResult;

  end;

procedure LancerAdminAgricole;

implementation
{$R *.dfm}

const
  cImgAdminDetailFermee     = 0;
  cImgAdminDetailOuvert     = 1;

  NbreNiveau  : Array [0..5] of Integer =(5,6,5,13,6,4);
  Niv1Detail  : Array [1..5] of string =('Culture','Cultures pérennes','Animaux','Transformation','Autres activités');
  Niv21Detail : Array [1..6] of string =('Céréales','Oléo-protéagineux','Fourrages','Cultures légumières','Cultures florales','Autres cultures');
  Niv22Detail : Array [1..5] of string =('Cultures fruitières','Viticulture','Sylviculture','Pépinières','Autres cultures pérennes');
  Niv23Detail : Array [1..13] of string =('Equins','Caprins','Ovins','Bovins','','','Porcs','Aviculture de chair','Avicultures oeufs','Gibiers','Coquillages','Ferme aquacole','Autres animaux');
  Niv24Detail : Array [1..6] of String =('Production laitière','Production boucherie','Gavages','Production boisson','Production fruitières','Autres transformations');
  Niv25Detail : Array [1..4] of String =('Autres activités végétales','Stockage','Tourisme rural','Prestations de services');


//--------------------------------
//--- Nom : LancerAdminAgricole
//--------------------------------
procedure LancerAdminAgricole;
var FAdminAgricole : TFAdminAgricole;
    PP             : THPanel;
begin
 FAdminAgricole:= TFAdminAgricole.Create(Appli);
 PP := FindInsidePanel;
 if PP = nil then
  begin
   try
    FAdminAgricole.ShowModal;
   finally
    FAdminAgricole.Free;
   end;
  end
 else
  begin
   InitInside (FAdminAgricole, PP);
   FAdminAgricole.Show;
   FAdminAgricole.SetFocus;
  end;
end;

//------------------------------------
//--- Nom : FormShow
//------------------------------------
procedure TFAdminAgricole.FormShow(Sender: TObject);
begin
 TV.Items.Clear;
 AlimenterTVAgricole;
end;

//------------------------------------
//--- Nom : AlimenterTVGed
//------------------------------------
procedure TFAdminAgricole.AlimenterTVAgricole;
var IndiceNiv1,IndiceNiv2,IndiceNiv3 : Integer;
    TAN1,TAN2,TAN3                   : TAgricoleNode;
    TN1, TN2, TN3                    : TTreeNode;
    SSql                             : String;
    UneTob                           : Tob;
begin
 //--- Traitement du niveau1
 for IndiceNiv1:= 1 to NbreNiveau [0] do
  begin
   TAN1 := TAgricoleNode.Create;
   TAN1.Niveau  := 1;
   TAN1.Code    := 100*IndiceNiv1;
   TAN1.Libelle := Niv1Detail [IndiceNiv1];
   TAN1.IconId  := cImgAdminDetailFermee;
   TAN1.IconSel := cImgAdminDetailOuvert;

   TN1 := TV.Items.AddChildObject(nil, TAN1.Libelle, TAN1);
   TN1.ImageIndex := TAN1.IconId;
   TN1.SelectedIndex := TAN1.IconSel;

   //--- Traitement du niveau 2
   for IndiceNiv2:= 1 to NbreNiveau [IndiceNiv1] do
    begin
     TAN2 := TAgricoleNode.Create;
     TAN2.Niveau := 2;
     TAN2.Code   := TAN1.Code+IndiceNiv2;
     if IndiceNiv1=1 then TAN2.Libelle := Niv21Detail [IndiceNiv2];
     if IndiceNiv1=2 then TAN2.Libelle := Niv22Detail [IndiceNiv2];
     if IndiceNiv1=3 then TAN2.Libelle := Niv23Detail [IndiceNiv2];
     if IndiceNiv1=4 then TAN2.Libelle := Niv24Detail [IndiceNiv2];
     if IndiceNiv1=5 then TAN2.Libelle := Niv25Detail [IndiceNiv2];

     if (TAN2.Libelle<>'') then
      begin
       TAN2.IconId  := cImgAdminDetailFermee;
       TAN2.IconSel := cImgAdminDetailOuvert;

       TN2 := TV.Items.AddChildObject(TN1, TAN2.Libelle, TAN2);
       TN2.ImageIndex := TAN2.IconId;
       TN2.SelectedIndex := TAN2.IconSel;

       //--- Traitement du niveau 3
       SSql:='SELECT YDS_CODE,YDS_LIBELLE FROM CHOIXDPSTD WHERE YDS_PREDEFINI="STD" AND YDS_NODOSSIER="000000" AND YDS_TYPE="DAG" AND YDS_CODE LIKE "'+IntToStr (TAN2.Code)+'%"';
       UneTob := Tob.Create('Agricole niveau 3', nil, -1);
       UneTob.LoadDetailFromSQL(SSql);
       for IndiceNiv3:=0 to UneTob.Detail.Count-1 do
        begin
         TAN3 := TAgricoleNode.Create;
         TAN3.Niveau  := 3;
         TAN3.Code    := UneTob.Detail[IndiceNiv3].GetValue('YDS_CODE');
         TAN3.Libelle := UneTob.Detail[IndiceNiv3].GetValue('YDS_LIBELLE');
         TAN3.IconId  := cImgAdminDetailFermee;
         TAN3.IconSel := cImgAdminDetailOuvert;

         TN3 := TV.Items.AddChildObject(TN2, TAN3.Libelle, TAN3);
         TN3.ImageIndex := TAN3.IconId;
         TN3.SelectedIndex := TAN3.IconSel;
        end;
       UneTob.Free;
      end
     else
      TAN2.free;
    end;
  end;

 //--- Se positionne sur le premier Noeud
 TN1 := TV.Items.GetFirstNode;
 if TN1<>Nil then TN1.Selected := True;
end;

//------------------------------------
//--- Nom : BInsertClick
//------------------------------------
procedure TFAdminAgricole.BinsertClick(Sender: TObject);
var UnLibelle     : string;
    MaxCode       : Integer;
    TN, TN3, Node : TTreeNode;
    TAN3          : TAgricoleNode;
begin
 TN:=TV.Selected;
 if TN=nil then Exit;
 if TN.Data=nil then exit;
 if (TAgricoleNode(TN.Data).Niveau <> 2) then exit;

 //--- Recherche du Maxcode
 Node := TN.getFirstChild;
 if Node <> Nil then
  begin
   MaxCode := TAgricoleNode (Node.Data).Code;
   while Node <> nil do
    begin
     if TAgricoleNode (Node.Data).Code > MaxCode then
      MaxCode := TAgricoleNode (Node.Data).Code;
     Node:=Node.getNextSibling;
    end;
  end
 else
  MaxCode:=TAgricoleNode (TN.Data).Code*1000;

 //--- Saisie et création du nouveau détail
 UnLibelle:='';
 if DonnerNomDetail (UnLibelle)=mrOk then
  begin
   TAN3:=TAgricoleNode.Create;
   TAN3.Niveau  := 3;
   TAN3.Code    := MaxCode+1;
   TAN3.Libelle := UnLibelle;
   TAN3.IconId  := cImgAdminDetailFermee;
   TAN3.IconSel := cImgAdminDetailOuvert;

   TN3:=TV.Items.AddChildObject(TN, TAN3.Libelle, TAN3);
   TN3.ImageIndex:=TAN3.IconId;
   TN3.SelectedIndex:=TAN3.IconSel;
   TN3.Selected:=True;
  end;

end;

//------------------------------------
//--- Nom : BDeleteClick
//------------------------------------
procedure TFAdminAgricole.BDeleteClick(Sender: TObject);
var TN : TTreeNode;
begin
 TN:=TV.Selected;
 if TN=nil then exit;
 if TN.Data=nil then exit;
 if TAgricoleNode(TN.Data).Niveau = 1 then exit;
 if TAgricoleNode(TN.Data).Niveau = 2 then exit;

 if UtilisationDossier (TAgricoleNode(TN.Data).Code) then
  begin
   PGIInfo('Impossible de supprimer '+TAgricoleNode (TN.Data).Libelle+' : il est utilisé dans un dossier.', 'Suppression d''un détail');
   Exit;
  end;

 if PGIAsk('Voulez-vous supprimer le détail '''+TAgricoleNode (TN.Data).Libelle+''' ?', 'Suppression d''un détail') = mrYes then
  TN.Delete;
end;

//------------------------------------
//--- Nom : UtilisationDossier
//------------------------------------
function TFAdminAgricole.UtilisationDossier (UnCode : Integer):Boolean;
begin
 Result := ExisteSQL('SELECT YDS_NODOSSIER FROM CHOIXDPSTD WHERE YDS_PREDEFINI="DOS" AND YDS_TYPE="DAG" AND YDS_CODE="'+IntToStr (UnCode)+'"');
end;

//------------------------------------
//--- Nom : FormCloseQuery
//------------------------------------
procedure TFAdminAgricole.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
 if PGIAsk('Voulez-vous enregistrer les modifications ?')=mrYes then
  BValider.Click;
end;

//------------------------------------
//--- Nom : DonnerNomDetail
//------------------------------------
function TFAdminAgricole.DonnerNomDetail (var Libelle : String) : TModalResult;
var Dialog : TFVierge;
    LHNomDetail : THLabel;
    THNomDetail : THEdit;
begin
 Dialog:=TFVierge.Create (Application);

 Dialog.Caption:='Saisie nom du détail';
 Dialog.Height:=120;
 Dialog.Width :=300;

 LHNomDetail:=THLabel.Create(Tcomponent (Dialog));
 LHNomDetail.Parent:=Dialog;
 LHNomDetail.Caption:='Nom';
 LHNomDetail.Top:=20;
 LHNomDetail.Left:=15;
 LHNomDetail.Height:=18;
 LHNomDetail.Width:=90;

 THNomDetail:=THEdit.Create(Tcomponent (Dialog));
 THNomDetail.Parent:=Dialog;
 THNomDetail.Text:=Libelle;
 THNomDetail.Top:=15;
 THNomDetail.Left:=120;
 THNomDetail.Height:=18;
 THNomDetail.Width:=150;

 Dialog.ShowModal;
 Libelle:=THNomDetail.Text;
 Result:=Dialog.ModalResult;
 Dialog.free;
end;


//------------------------------------
//--- Nom : BValiderClick
//------------------------------------
procedure TFAdminAgricole.BValiderClick(Sender: TObject);
Var TN : TTreeNode;
begin
 inherited;
 //--- Suppression de tout les détails
 ExecuteSQL ('DELETE FROM CHOIXDPSTD WHERE YDS_PREDEFINI="STD" AND YDS_NODOSSIER="000000" AND YDS_TYPE="DAG"');

 //--- Sauvegarde de l'ensemble de l'arbre niveau3
 TN:=TV.Items.GetFirstNode;
 while (TN <> nil) do
  begin
   if (TagricoleNode (TN.Data).Niveau=3) then
    ExecuteSQl ('INSERT INTO CHOIXDPSTD (YDS_PREDEFINI,YDS_NODOSSIER,YDS_TYPE,YDS_CODE,YDS_LIBELLE) values ("STD","000000","DAG","'+IntToStr (TagricoleNode (TN.Data).Code)+'","'+TagricoleNode (TN.Data).Libelle+'")');
   TN:= TN.GetNext;
  end;
end;


end.
