{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 25/10/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
unit UOutilsAnnuLienTri;
//////////////////////////////////////////////////////////////////
interface
//////////////////////////////////////////////////////////////////
uses
   UComOutils, Classes, Hctrls, SysUtils, UTOB;

//////////////////////////////////////////////////////////////////
procedure ReChargeCurLiens(sGuidPerDos_p, sNoOrdre_p,
                           sTypeDos_p, sTypeFiche_p, sForme_p, sFonction_p : string;
                           var OBFonctions_p : TOB);

procedure ChargeFonctionsEtLiens(sGuidPerDos_p, sNoOrdre_p,
                          sTypeDos_p, sTypeFiche_p, sForme_p : string;
                          var OBFonctions_p, OBCurListeLiens_p, OBCurLien_p : TOB);

function  ChargeLiens(sGuidPerDos_p, sNoOrdre_p, sTypeDos_p, sTypeFiche_p, sUneFonction_p : string) : TOB;
function  ChargeFonctions(sForme_p, sTypeDos_p, sTypeFiche_p : string) : TOB;

procedure RegroupeTOBLien(OBFonctions_p : TOB; var OBLiens_p : TOB);
procedure CalculNbLienTOB(OBFonctions_p : TOB);

{procedure ChargeEtAfficheFonctionsEtLiens(sGuidPerDos_p, sNoOrdre_p,
                          sTypeDos_p, sTypeFiche_p, sForme_p : string;
                          GListeFonction_p, GListeLien_p : THGrid;
                          var OBFonctions_p, OBCurListeLiens_p, OBCurLien_p : TOB);}

procedure AfficheFonctionsEtLiens(sTypeDos_p : string;
                          GListeFonction_p, GListeLien_p : THGrid;
                          var OBFonctions_p, OBCurListeLiens_p, OBCurLien_p : TOB);

procedure AfficheCurLiens(GListeFonction_p, GListeLien_p : THGrid;
                          var OBCurListeLiens_p, OBCurLien_p : TOB);

function SelectCurListeLiens(GListeFonction_p : THGrid) : TOB;
function SelectCurLien(GListeLien_p : THGrid) : TOB;

procedure FormateListes(GListeFonction_p, GListeLien_p : THGrid);

//////////////////////////////////////////////////////////////////
implementation
//////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/04/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure ChargeEtAfficheFonctionsEtLiens(sGuidPerDos_p, sNoOrdre_p,
                          sTypeDos_p, sTypeFiche_p, sForme_p : string;
                          GListeFonction_p, GListeLien_p : THGrid;
                          var OBFonctions_p, OBCurListeLiens_p, OBCurLien_p : TOB);
var
   OBLiens_l : TOB;
begin

   ChargeFonctionsEtLiens(sGuidPerDos_p, sNoOrdre_p,
                          sTypeDos_p, sTypeFiche_p, sForme_p,
                          OBFonctions_p, OBCurListeLiens_p, OBCurLien_p);

   AfficheFonctionsEtLiens(sTypeDos_p, GListeFonction_p, GListeLien_p,
                             OBFonctions_p, OBCurListeLiens_p, OBCurLien_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 07/02/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure ChargeFonctionsEtLiens(sGuidPerDos_p, sNoOrdre_p,
                          sTypeDos_p, sTypeFiche_p, sForme_p : string;
                          var OBFonctions_p, OBCurListeLiens_p, OBCurLien_p : TOB);
var
   OBLiens_l : TOB;
begin
   if OBFonctions_p <> nil then
      FreeAndNil(OBFonctions_p);

   OBFonctions_p := ChargeFonctions(sForme_p, sTypeDos_p, sTypeFiche_p);
   OBLiens_l := ChargeLiens(sGuidPerDos_p, sNoOrdre_p, sTypeDos_p, sTypeFiche_p, '');

   RegroupeTOBLien(OBFonctions_p, OBLiens_l);
   CalculNbLienTOB(OBFonctions_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 07/02/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure AfficheFonctionsEtLiens(sTypeDos_p : string;
                          GListeFonction_p, GListeLien_p : THGrid;
                          var OBFonctions_p, OBCurListeLiens_p, OBCurLien_p : TOB);
var
   iOBInd_l, iGDRow_l : integer;
begin
   FormateListes(GListeFonction_p, GListeLien_p);
//   OBFonctions_p.PutGridDetail(GListeFonction_p, false, TRUE, 'JTF_FONCTABREGE;TOTAL', true);
   iGDRow_l := 0;
   for iOBInd_l := 0 to OBFonctions_p.Detail.Count - 1 do
   begin
      if (sTypeDos_p <> '') and (sTypeDos_p <> OBFonctions_p.Detail[iOBInd_l].GetString('JFT_TYPEDOS')) then
         Continue;
      Inc(iGDRow_l);
      GListeFonction_p.RowCount := iGDRow_l + 1;
      OBFonctions_p.Detail[iOBInd_l].PutLigneGrid(GListeFonction_p, iGDRow_l, false, TRUE, 'JTF_FONCTABREGE;TOTAL', true);
   end;

   AfficheCurLiens(GListeFonction_p, GListeLien_p, OBCurListeLiens_p, OBCurLien_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/04/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure FormateListes(GListeFonction_p, GListeLien_p : THGrid);
begin
    with GListeFonction_p do
    begin
       TitleCenter := true;
       TitleBold := true;
       ColAligns[1] := taRightJustify;
       Cells[0, 0] := 'Fonctions';
       Cells[1, 0] := 'Total';
//       Row := 1;
    end;

    with GListeLien_p do
    begin
       TitleCenter := true;
       TitleBold := true;
       Cells[0, 0] :='Code';
       Cells[1, 0] :='Nom';
       Cells[2, 0] :='Info';
       ColAligns[2] := taRightJustify;
//       Row := 1;
    end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/04/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure ReChargeCurLiens(sGuidPerDos_p, sNoOrdre_p,
                           sTypeDos_p, sTypeFiche_p, sForme_p, sFonction_p : string;
                           var OBFonctions_p  : TOB);
var
   OBFonction_l, OBLiens_l : TOB;
begin
   OBFonction_l := OBFonctions_p.FindFirst(['JTF_FONCTION'], [sFonction_p], FALSE) ;
   OBFonction_l.ClearDetail;

   OBLiens_l := ChargeLiens(sGuidPerDos_p, sNoOrdre_p, sTypeDos_p, sTypeFiche_p, sFonction_p);
   RegroupeTOBLien(OBFonctions_p, OBLiens_l);
   CalculNbLienTOB(OBFonctions_p);
   OBLiens_l.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/04/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function ChargeLiens(sGuidPerDos_p, sNoOrdre_p, sTypeDos_p, sTypeFiche_p, sUneFonction_p : string) : TOB;
var
   OBLiens_l : TOB;
   sReq_l : string;
   iInd_l : integer;
begin
   sReq_l := 'select * from ANNULIEN ' +
             'where ANL_GUIDPERDOS = "' + sGuidPerDos_p + '" ' ; 

   if sNoOrdre_p <> '' then
      sReq_l := sReq_l + '  AND ANL_NOORDRE = ' + sNoOrdre_p + ' ' ;

   if sTypeDos_p <> '' then
      sReq_l := sReq_l + '  AND ANL_TYPEDOS = "' + sTypeDos_p + '" ';

   if sUneFonction_p <> '' then
   begin
      sReq_l := sReq_l +  '  AND ANL_FONCTION = "' + sUneFonction_p + '"';
   end
   else
   begin
      if sTypeDos_p <> '' then
         sReq_l := sReq_l +  '  AND ANL_FONCTION <> "' + sTypeDos_p + '"';

      if sTypeFiche_p = 'DP' then
         sReq_l := sReq_l + '  AND ANL_FONCTION <> "INT"';
   end;

   sReq_l := sReq_l +  ' order by ANL_TRI, ANL_NOMPER';

   OBLiens_l := TOB.Create('ANNULIEN', NIL, -1);
   OBLiens_l.LoadDetailDBFromSQL('ANNULIEN', sReq_l);

   for iInd_l := 0 to OBLiens_l.Detail.count - 1 do
   begin
       OBLiens_l.Detail[iInd_l].AddChampSup('CALCUL', true);
       OBLiens_l.Detail[iInd_l].PutValue('CALCUL', 1);
       OBLiens_l.Detail[iInd_l].AddChampSup('FONCTION', true);
       OBLiens_l.Detail[iInd_l].PutValue('FONCTION', OBLiens_l.Detail[iInd_l].GetValue('ANL_FONCTION'));
       OBLiens_l.Detail[iInd_l].AddChampSup('TITRE', true);
       OBLiens_l.Detail[iInd_l].PutValue('TITRE', '-');
   end;

   result := OBLiens_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/04/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function ChargeFonctions(sForme_p, sTypeDos_p, sTypeFiche_p : string) : TOB;
var
   iInd_l : integer;
   sReq_l : string;
   OBFonctions_l : TOB;
begin
   sReq_l := 'select JTF_FONCTION, JTF_FONCTABREGE, JFT_TRI, JFT_TITRE, JFT_TYPEDOS ' +
             'from JUTYPEFONCT left join JUFONCTION on JFT_FONCTION = JTF_FONCTION ';

   if sTypeDos_p <> '' then
   begin
      sReq_l := sReq_l + 'where JFT_TYPEDOS = "' + sTypeDos_p + '" ';
      if sTypeDos_p = 'STE' then
         sReq_l := sReq_l + '  AND JFT_FORME = "' + sForme_p + '" ';
   end
   else
   begin
      sReq_l := sReq_l + 'where ((JFT_TYPEDOS = "STE" AND JFT_FORME = "' + sForme_p + '") ' +
                         '   OR (JFT_TYPEDOS <> "STE")) ';
   end;

   sReq_l := sReq_l + '  AND (JTF_FONCTION <> "STE") ';
   if sTypeFiche_p = 'DP' then
      sReq_l := sReq_l + '  AND (JTF_FONCTION <> "INT") ';

   sReq_l := sReq_l + 'ORDER BY JFT_TRI';

   OBFonctions_l := TOB.Create('JUTYPEFONCT', NIL, -1);
   OBFonctions_l.LoadDetailDBFromSQL('JUTYPEFONCT', sReq_l);
   for iInd_l := 0 to OBFonctions_l.Detail.count-1 do
   begin
      OBFonctions_l.Detail[iInd_l].AddChampSup('TOTAL', true);
      OBFonctions_l.Detail[iInd_l].PutValue('TOTAL', 0);
   end;
   result := OBFonctions_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure RegroupeTOBLien(OBFonctions_p : TOB; var OBLiens_p : TOB);
var
  iInd_l : integer;
  bFin_l : boolean;
  sFonction_l : string;
  OBFonct_l : TOB ;
begin
   // ON BOUCLE SUR la liste des liens POUR AFFECTER LA TOB A LA FONCTION
   if OBLiens_p.Detail.Count >= 1 then
   begin
      iInd_l := 0 ;
      bFin_l := FALSE ;
      while not bFin_l do
      begin
         sFonction_l := OBLiens_p.Detail[iInd_l].GetValue('ANL_FONCTION');
         OBFonct_l := OBFonctions_p.FindFirst(['JTF_FONCTION'], [sFonction_l], FALSE) ;
         if OBFonct_l <> nil then
         begin
            OBLiens_p.Detail[iInd_l].PutValue('TITRE', OBFonct_l.GetString('JFT_TITRE'));
            OBLiens_p.Detail[iInd_l].ChangeParent(OBFonct_l, -1) ;

            // BM 20/08/2002 = I sinon indice hors limite après incrémentation de I
            // car boucle une fois de trop
            if OBLiens_p.Detail.Count = iInd_l then
               bFin_l := TRUE;
         end
         else
         begin
            Inc(iInd_l) ;
            if OBLiens_p.Detail.Count = iInd_l then
               bFin_l := TRUE;
         end;
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure CalculNbLienTOB(OBFonctions_p : TOB);
var
  iInd_l : INTEGER ;
  sFonction_l : string;
  dCalcul_l : double;
  sCalcul_l : string;
begin
   for iInd_l := 0 to OBFonctions_p.Detail.count - 1 do
   begin
      sFonction_l := OBFonctions_p.Detail[iInd_l].GetValue('JTF_FONCTION');
      sCalcul_l := '0';
      if OBFonctions_p.Detail[iInd_l].Detail.Count >= 1 then
      begin
          dCalcul_l := OBFonctions_p.Detail[iInd_l].Somme('CALCUL', ['ANL_FONCTION'], [sFonction_l], true);
          sCalcul_l := FloatToStr(dCalcul_l);
      end;
      OBFonctions_p.Detail[iInd_l].PutValue('TOTAL', sCalcul_l);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/04/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure AfficheCurLiens(GListeFonction_p, GListeLien_p : THGrid;
                          var OBCurListeLiens_p, OBCurLien_p : TOB);
begin
   OBCurListeLiens_p := SelectCurListeLiens(GListeFonction_p);

   if OBCurListeLiens_p = nil then
   begin
      GListeLien_p.VidePile(false);
   end
   else
   begin
      OBCurListeLiens_p.PutGridDetail(GListeLien_p, FALSE, TRUE, 'ANL_AFFICHE;ANL_NOMPER;ANL_INFO', TRUE) ;
      GListeLien_p.SortGrid(1, FALSE);

      LargeurUneColonne(GListeLien_p, 0);
      LargeurUneColonne(GListeLien_p, 2);

      OBCurLien_p := SelectCurLien(GListeLien_p);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function SelectCurListeLiens(GListeFonction_p : THGrid) : TOB;
begin
   result := nil;
   if (GListeFonction_p.Row > 0) and (GListeFonction_p.Row < GListeFonction_p.RowCount) then
      result := TOB(GListeFonction_p.Objects[0, GListeFonction_p.Row]) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function SelectCurLien(GListeLien_p : THGrid) : TOB;
begin
   result := nil;
   if (GListeLien_p.Row > 0) and (GListeLien_p.Row <GListeLien_p.RowCount) then
      result := TOB(GListeLien_p.Objects[0, GListeLien_p.Row]) ;
end;
//////////////////////////////////////////////////////////////////
end.
