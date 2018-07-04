unit UFNewForme;

interface

uses
   Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
   Vierge, HSysMenu, HTB97, Grids, DBGrids, HDB, ComCtrls, StdCtrls, Hctrls,
   ExtCtrls, Db, DBTables, HEnt1, HStatus, UiUtil, HPanel, UTOB,
   ULanceFiche, HMsgBox;

type
   TFNewForme = class(TFVierge)
      PForme_f: TPanel;
      LNewForme_f: TLabel;
    eForme_f: TEdit;
    eFormeLib_f: TEdit;
      Pfonction_f: TPanel;
      GFonction_f: THGrid;
      PBT_f: TPanel;
      BFoncToType_f: TToolbarButton97;
      BTypeToFonc_f: TToolbarButton97;
      GTypeFonct_f: THGrid;


      procedure FormShow(Sender: TObject);
      procedure FormClose(Sender: TObject; var Action: TCloseAction);

      procedure BValiderClick(Sender: TObject);
      procedure BFoncToType_fClick(Sender: TObject);
      procedure BTypeToFonc_fClick(Sender: TObject);

  private
      sForme_c, sTypeDos_c : string;

      OBFonction_c, OBFoncDel_c, OBTypeFonct_c, OBTypeLess_c : TOB;

      procedure ChargeTypeFonct;
      procedure ChargeFonction;

      procedure AfficheTypeFonct;
      procedure AfficheFonction;

      procedure FormateGrille(GGrid_p : THGrid; aosTitre_p : array of string);

      procedure EnleveTypeFonct;
      procedure Deplace(GFrom_p : THGrid; sFrom_p : string);
      procedure TypeToFonc(sFonction_p : string);
      procedure FoncToType(sFonction_p : string);

      function FindFoncLib(sFonction_p : string) : string;      
  public

  end;

/////////////////////////////////////////////////////////////////
procedure NewForme(sForme_p, sFormeLib_p, sTypeDos_p, sBase_p : string);
procedure DelForme(sForme_p, sTypeDos_p : string);

/////////////////////////////////////////////////////////////////
implementation

{$R *.DFM}
/////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure NewForme(sForme_p, sFormeLib_p, sTypeDos_p, sBase_p : string);
var
   FNewForme : TFNewForme;
begin
   FNewForme := TFNewForme.Create(Application);
   FNewForme.sForme_c := sForme_p;
   FNewForme.sTypeDos_c := sTypeDos_p;
   FNewForme.eForme_f.Text := sForme_p;
   FNewForme.eFormeLib_f.Text := sFormeLib_p;
   LanceLaFiche(FNewForme);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/07/2005
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure DelForme(sForme_p, sTypeDos_p : string);
begin
//   if PGIAsk('Voulez-vous supprimer les fonctions associées' + #13#10 +
//             'à cette forme juridique?', 'Forme juridique ' + sForme_p) = mrYes then
      ExecuteSQL('delete from jufonction where jft_forme = "' + sForme_p + '"');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/07/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFNewForme.FormShow(Sender: TObject);
begin
   inherited;

   BValider.Enabled := false;
   OBFonction_c := TOB.Create('JUFONCTION', nil, -1);
   OBFoncDel_c := TOB.Create('JUFONCTION', nil, -1);
   OBTypeFonct_c := TOB.Create('JUTYPEFONCT', nil, -1);
   OBTypeLess_c := TOB.Create('JUTYPEFONCT', nil, -1);

   ChargeTypeFonct;
   ChargeFonction;
   EnleveTypeFonct;
   AfficheTypeFonct;
   AfficheFonction;
   BValider.Enabled := (OBFonction_c.Detail.Count > 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/07/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFNewForme.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   inherited;
   OBFonction_c.Free;
   OBFoncDel_c.Free;
   OBTypeFonct_c.Free;
   OBTypeLess_c.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 27/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFNewForme.ChargeTypeFonct;
begin
   OBTypeFonct_c.ClearDetail;
   OBTypeFonct_c.LoadDetailDBFromSQL('JUTYPEFONCT', 'SELECT * FROM JUTYPEFONCT ORDER BY JTF_FONCTION');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 27/07/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFNewForme.ChargeFonction;
begin
   OBFonction_c.ClearDetail;
   OBFonction_c.LoadDetailDBFromSQL('JUFONCTION',
         'SELECT * FROM JUFONCTION ' +
         'WHERE JFT_TYPEDOS = "' + sTypeDos_c + '" AND JFT_FORME = "' + sForme_c + '" ' +
         'ORDER BY JFT_FONCTION');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 27/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFNewForme.AfficheTypeFonct;
begin
   GTypeFonct_f.VidePile(false);
   OBTypeFonct_c.PutGridDetail(GTypeFonct_f, true, true, 'JTF_FONCTION;JTF_FONCTABREGE');
   FormateGrille(GTypeFonct_f, ['Fonction', 'Libellé']);
   GTypeFonct_f.SortGrid(0, false);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 27/07/2005
Modifié le ... : 27/07/2005
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFNewForme.AfficheFonction;
var
   iInd_l : integer;
   sFonction_l : string;
begin
   GFonction_f.VidePile(false);
   OBFonction_c.PutGridDetail(GFonction_f, true, true, 'JFT_FORME;JFT_FONCTION;JTF_FONCTABREGE');

   for iInd_l := 1 to GFonction_f.RowCount - 1 do
   begin
      sFonction_l := GFonction_f.Cells[1, iInd_l];
      if sFonction_l <> '' then
         GFonction_f.Cells[2, iInd_l] := FindFoncLib(sFonction_l);
   end;

   FormateGrille(GFonction_f, ['Forme', 'Fonction', 'Libellé']);
   GFonction_f.SortGrid(1, false);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 27/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFNewForme.EnleveTypeFonct;
var
   OBDetail_l : TOB;
   sFonction_l : string;
   iInd_l : integer;
begin
   OBTypeLess_c.ClearDetail;
   OBFoncDel_c.ClearDetail;
   for iInd_l := 0 to OBFonction_c.Detail.Count - 1 do
   begin
      sFonction_l := OBFonction_c.Detail[iInd_l].GetString('JFT_FONCTION');
      OBDetail_l := OBTypeFonct_c.FindFirst(['JTF_FONCTION'], [sFonction_l], true);
      if OBDetail_l <> nil then
         OBDetail_l.ChangeParent(OBTypeLess_c, -1);
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 27/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFNewForme.BFoncToType_fClick(Sender: TObject);
begin
   inherited;
   Deplace(GFonction_f, 'FoncToType');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 27/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFNewForme.BTypeToFonc_fClick(Sender: TObject);
begin
  inherited;
   Deplace(GTypeFonct_f, 'TypeToFonc');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 27/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFNewForme.Deplace(GFrom_p : THGrid; sFrom_p : string);
var
   iInd_l, iRow_l : integer;
begin
   inherited;
   if GFrom_p.AllSelected then
   begin
      for iRow_l := GFrom_p.RowCount - 1 downto 1 do
      begin
         if sFrom_p = 'FoncToType' then
            FoncToType(GFrom_p.Cells[2, iRow_l])
         else
            TypeToFonc(GFrom_p.Cells[0, iRow_l]);
      end;
        GFrom_p.AllSelected := False;
   end
   else if GFrom_p.NbSelected <> 0 then
   begin
      for iInd_l := GFrom_p.NbSelected - 1 downto 0 do
      begin
         GFrom_p.GotoLeBookmark(iInd_l);
         iRow_l := GFrom_p.Row;
         if sFrom_p = 'FoncToType' then
            FoncToType(GFrom_p.Cells[2, iRow_l])
         else
            TypeToFonc(GFrom_p.Cells[0, iRow_l]);
      end;
      GFrom_p.ClearSelected;
   end
   else
      Exit;

   AfficheFonction;
   AfficheTypeFonct;

   BFoncToType_f.Enabled := (OBFonction_c.Detail.Count > 0);
   BTypeToFonc_f.Enabled := (OBTypeFonct_c.Detail.Count > 0);
   BValider.Enabled := (OBFonction_c.Detail.Count > 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 27/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFNewForme.TypeToFonc(sFonction_p : string);
var
   OBDetailFonc_l, OBDetailType_l : TOB;
begin
   OBDetailType_l := OBTypeFonct_c.FindFirst(['JTF_FONCTION'], [sFonction_p], true);
   if OBDetailType_l <> nil then
      OBDetailType_l.ChangeParent(OBTypeLess_c, -1);

   OBDetailFonc_l := OBFoncDel_c.FindFirst(['JFT_FONCTION'], [sFonction_p], true);
   if OBDetailFonc_l <> nil then
      OBDetailFonc_l.ChangeParent(OBFonction_c, -1)
   else
      OBDetailFonc_l := TOB.Create('JUFONCTION', OBFonction_c, -1);

   OBDetailFonc_l.PutValue('JFT_TYPEDOS', sTypeDos_c);
   OBDetailFonc_l.Putvalue('JFT_FORME', sForme_c);
   OBDetailFonc_l.PutValue('JFT_FONCTION', sFonction_p);
   OBDetailFonc_l.PutValue('JFT_APPMODIF', OBDetailType_l.GetString('JTF_APPMODIF'));
   OBDetailFonc_l.PutValue('JFT_DATECREATION', Date);
   OBDetailFonc_l.PutValue('JFT_DATEMODIF', Date);
   OBDetailFonc_l.PutValue('JFT_PREDEFINI', OBDetailType_l.GetString('JTF_PREDEFINI'));
   OBDetailFonc_l.PutValue('JFT_RACINE', OBDetailType_l.GetString('JTF_RACINE'));
   OBDetailFonc_l.PutValue('JFT_TIERS', OBDetailType_l.GetString('JTF_TIERS'));
   OBDetailFonc_l.PutValue('JFT_TITRE', OBDetailType_l.GetString('JTF_TITRE'));
   OBDetailFonc_l.PutValue('JFT_UTILISATEUR', OBDetailType_l.GetString('JTF_UTILISATEUR'));
   OBDetailFonc_l.PutValue('JFT_VERSION', 9);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 27/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TFNewForme.FoncToType(sFonction_p : string);
var
   OBDetailFonc_l, OBDetailType_l : TOB;
begin
   OBDetailFonc_l := OBFonction_c.FindFirst(['JFT_FONCTION'], [sFonction_p], true);
   if OBDetailFonc_l <> nil then
      OBDetailFonc_l.ChangeParent(OBFoncDel_c, -1);

   OBDetailType_l := OBTypeLess_c.FindFirst(['JTF_FONCTION'], [sFonction_p], true);
   if OBDetailType_l <> nil then
      OBDetailType_l.ChangeParent(OBTypeFonct_c, -1);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 27/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFNewForme.FormateGrille(GGrid_p : THGrid; aosTitre_p : array of string);
var
   iColInd_l, iRowInd_l, iWidth_l : integer;
   aoiWidth_l : array of integer;
begin
   GGrid_p.ColCount := Length(aosTitre_p);
   For iColInd_l := 0 to Length(aosTitre_p) - 1 do
      GGrid_p.Cells[iColInd_l, 0] := aosTitre_p[iColInd_l];

   SetLength(aoiWidth_l, GGrid_p.ColCount);
   for iColInd_l := 0 to GGrid_p.ColCount - 1 do
      aoiWidth_l[iColInd_l] := GGrid_p.Canvas.TextWidth(GGrid_p.Cells[ iColInd_l, 0 ]) + 15;

   for iRowInd_l := 1 to GGrid_p.RowCount - 1 do
   begin
      for iColInd_l := 0 to GGrid_p.ColCount - 1 do
      begin
         iWidth_l := GGrid_p.Canvas.TextWidth(GGrid_p.Cells[iColInd_l, iRowInd_l ]) + 15;
         if aoiWidth_l[iColInd_l] < iWidth_l then
            aoiWidth_l[iColInd_l] := iWidth_l;
        end;
     end;

   for iColInd_l := 0 to GGrid_p.ColCount - 1 do
        GGrid_p.ColWidths[iColInd_l] := aoiWidth_l[iColInd_l];

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 27/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TFNewForme.FindFoncLib(sFonction_p : string) : string;
var
   OBDetail_l : TOB;
begin
   OBDetail_l := OBTypeFonct_c.FindFirst(['JTF_FONCTION'], [sFonction_p], true);
   if OBDetail_l = nil then
      OBDetail_l := OBTypeLess_c.FindFirst(['JTF_FONCTION'], [sFonction_p], true);
   result := OBDetail_l.GetString('JTF_FONCTABREGE');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/07/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFNewForme.BValiderClick(Sender: TObject);
var
   iInd_l : integer;
begin
   inherited;
   for iInd_l := 0 to OBFoncDel_c.Detail.Count - 1 do
   begin
      if OBFoncDel_c.Detail[iInd_l].ExistDB then
         OBFoncDel_c.Detail[iInd_l].DeleteDB;
   end;

   OBFonction_c.InsertOrUpdateDB;
end;


end.
