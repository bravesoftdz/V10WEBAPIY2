{***********UNITE*************************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 23/06/2004
Modifié le ... :   /  /
Description .. : Chargement de données dans une grille non reliée à un MUL
Mots clefs ... :
*****************************************************************}
Unit CLASSGrilleSansMul ;

Interface

Uses
   HEnt1, sysutils, Classes, HCtrls, UTOB, hmsgbox, grids,

   {$IFDEF VER150}
   Variants,
   {$ENDIF}

   extctrls, Controls;

/////////////////////////////////////////////////////////////////
type
   TGrilleSansMul = class

      public
         constructor Create(gdGrille_p : THGrid; sListe_p : string;
                            PCumul_p : TPanel); overload;

         destructor  Destroy(bok_p : boolean=false); overload;

         procedure OnLoad(sWhere_p : string = '';
                         bAvecCpt_p : boolean = false; iNbLignes_p : integer = -1); overload;
         procedure OnLoad(sRequete_p, sTitres_p, sLargeurs_p, sFormats_p : string;
                         bAvecCpt_p : boolean = false; iNbLignes_p : integer = -1); overload;

         procedure OnReload(bAvecCpt_p : boolean = false; iNbLignes_p : integer = -1);

         function OnGetField(sChamp_p : string) : variant;

         function ListeCorrecte(sChamp_p, sLeTitre_p : String): Boolean;
         function ChampEstDansListe(sChamp_p : String): integer;

         procedure OnDblClick(Sender : TObject); overload;
         function NbRows : integer;
      private
         gdGrille_c  : THGrid;
         PCumul_c    : TPanel;

         bAvecCpt_c, bDepuisReq_c  : boolean;
         iNbLignes_c : integer;
         iCurRow_c   : integer;

         OBListe_c : TOB;
         OBTable_c : TOB;
      
         sListe_c    : string;
         sWhere_c    : string;
         sTable_c    : string;
         sChamps_c   : string;
         sJointure_c : string;
         sOrdre_c    : string;
         sTitres_c   : string;
         sLargeurs_c : string;
         sFormats_c  : string;

         function  ChargeLaGrille(sTable_p, sChamps_p, sWhere_p, sJointure_p,
                                  sOrdre_p, sFormats_p : string) : boolean; overload;

         function  ChargeLaGrille(sRequete_p, sAffChamps_p : string) : boolean; overload;


         function  ChargeLaListe(var tsMemo_p : TStrings) : boolean;

         procedure FormateLaGrille(sTitres_p, sLargeurs_p, sFormats_p : string);


         function  DonneesTablette(sTablette_p : string;
                                   var sCode_p, sChampLib_p, sTable_p, sWhere_p, sPrefixe_p : string) : boolean;

         function  ExtractChamps(sRequete_p : string) : string;

end;

/////////////////////////////////////////////////////////////////

//////////////////////// IMPLEMENTATION /////////////////////////

Implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 28/06/2004
Modifié le ... : 29/06/2004
Description .. : Chargement des données depuis la liste, affichage et
Suite ........ : formatage dans la grille
Mots clefs ... :
*****************************************************************}
constructor TGrilleSansMul.Create(gdGrille_p : THGrid; sListe_p : string;
                                  PCumul_p : TPanel);
begin
   gdGrille_c  := gdGrille_p;
   PCumul_c    := PCumul_p;
   sListe_c    := sListe_p;
   bAvecCpt_c  := false;
   iNbLignes_c := -1;

   OBListe_c := TOB.Create('LISTE', nil, -1);
   OBTable_c := TOB.Create('TABLE', nil, -1);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 12/08/2004
Modifié le ... : 16/08/2004
Description .. : 
Mots clefs ... : 
*****************************************************************}
destructor TGrilleSansMul.Destroy(bok_p : boolean=false);
begin
   if OBListe_c <> nil then
      FreeAndNil(OBListe_c);
   if OBTable_c <> nil then
      FreeAndNil(OBTable_c);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 12/08/2004
Modifié le ... : 13/08/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGrilleSansMul.OnLoad(sWhere_p : string = '';
                            bAvecCpt_p : boolean = false; iNbLignes_p : integer = -1);
var
   tsMemo_l : TStrings;
begin
   bAvecCpt_c := bAvecCpt_p;
   bDepuisReq_c := false;
   iNbLignes_c := iNbLignes_p;

   // Chargement du mémo
   tsMemo_l := TStringList.Create;
   if not ChargeLaListe(tsMemo_l) then
   begin
      tsMemo_l.Free;
      Exit;
   end;

   sTable_c    := UpperCase(Trim(tsMemo_l[0]));
   sChamps_c   := UpperCase(Trim(tsMemo_l[1]));
   sJointure_c := UpperCase(Trim(tsMemo_l[2]));
   sOrdre_c    := UpperCase(Trim(tsMemo_l[3]));
   sTitres_c   := Trim(tsMemo_l[4]);
   sLargeurs_c := UpperCase(Trim(tsMemo_l[5]));
   sFormats_c  := UpperCase(Trim(tsMemo_l[6]));

   sWhere_c    := sWhere_p;
   tsMemo_l.free;
   OnReload(bAvecCpt_p, iNbLignes_p);
end;

procedure TGrilleSansMul.OnLoad(sRequete_p, sTitres_p, sLargeurs_p, sFormats_p : string;
                                bAvecCpt_p : boolean = false; iNbLignes_p : integer = -1);
begin
   bAvecCpt_c := bAvecCpt_p;
   bDepuisReq_c := true;
   iNbLignes_c := iNbLignes_p;

   // Liste
   if (sTitres_p = '') or (sLargeurs_p = '') or (sFormats_p = '') then
      Exit;

   sTitres_c   := sTitres_p;
   sLargeurs_c := UpperCase(sLargeurs_p);
   sFormats_c  := UpperCase(sFormats_p);

   sWhere_c    := UpperCase(sRequete_p);
   sChamps_c   := ExtractChamps(sRequete_p);
   sChamps_c   := StringReplace(sChamps_c, 'DISTINCT ', '', [rfReplaceAll, rfIgnoreCase]);
   OnReload(bAvecCpt_p, iNbLignes_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 17/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TGrilleSansMul.OnReload(bAvecCpt_p : boolean = false; iNbLignes_p : integer = -1);
begin
   bAvecCpt_c := bAvecCpt_p;
   iNbLignes_c := iNbLignes_p;

   if bDepuisReq_c then
      ChargeLaGrille(sWhere_c, sChamps_c)
   else
      ChargeLaGrille(sTable_c, sChamps_c, sWhere_c, sJointure_c, sOrdre_c, sFormats_c);
   FormateLaGrille(sTitres_c, sLargeurs_c, sFormats_c);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 12/08/2004
Modifié le ... : 13/08/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGrilleSansMul.OnDblClick(Sender : TObject);
begin
    gdGrille_c.OnDblClick(Sender);
    iCurRow_c := gdGrille_c.Row;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 13/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGrilleSansMul.OnGetField(sChamp_p : string) : variant;
var
   iCol_l, iRow_l : integer;
begin
   iCol_l := ChampEstDansListe(sChamp_p);
   if iCol_l = -1 then
   begin
      PGIInfo('Champ "' + sChamp_p + '" inconnu dans la liste.', TitreHalley);
      result := Unassigned;
   end
   else
   begin
      iRow_l := gdGrille_c.Row;
      result := OBTable_c.Detail[iRow_l - 1].GetValue(sChamp_p);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 28/06/2004
Modifié le ... :   /  /
Description .. : Chargement de la liste
Mots clefs ... :
*****************************************************************}
function  TGrilleSansMul.ChargeLaListe(var tsMemo_p : TStrings) : boolean;
var
   sLesData_l, sData_l : string;

begin
   result := false;

   OBListe_c.LoadDetailDBFromSQL('LISTE', 'SELECT LI_DATA,LI_UTILISATEUR FROM LISTE ' +
                                 'WHERE LI_LISTE = "' + sListe_c + '" ' +
                                 '  AND LI_UTILISATEUR = "' + V_PGI.User + '"');

   if OBListe_c.Detail.Count = 0 then
   begin
      OBListe_c.ClearDetail;
      OBListe_c.LoadDetailDBFromSQL('LISTE', 'SELECT LI_DATA,LI_UTILISATEUR FROM LISTE ' +
                                    'WHERE LI_LISTE = "' + sListe_c + '" ' +
                                    '  AND LI_UTILISATEUR = "---"');
      if OBListe_c.Detail.Count = 0 then
         Exit;
   end;

   // Chargement du mémo
   sLesData_l := VarAsType(OBListe_c.Detail[0].GetValue('LI_DATA'), varString);
   while sLesData_l <> '' do
   begin
      sData_l := READTOKENPipe(sLesData_l, #$D#$A);
      tsMemo_p.Add(sData_l);
   end;

   if tsMemo_p.Count < 9 then
      exit;
   result := true;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 28/06/2004
Modifié le ... : 06/07/2004
Description .. : Chargement des données dans la grille par construction de
Suite ........ : la requête
Mots clefs ... : 
*****************************************************************}
function  TGrilleSansMul.ChargeLaGrille(sTable_p, sChamps_p, sWhere_p, sJointure_p,
                         sOrdre_p, sFormats_p : string) : boolean;
var
   sReqChamps_l, sReqTables_l, sRequete_l, sReqJoin_l : string;
   sLeChamp_l, sAffChamps_l, sFormat_l, sPrefixe_l : string;
   sCode_l, sChampLib_l, sTable_l, sWhere_l, sOrdre_l : string;
   iOrdre_l, iInd_l : integer;
begin
   result := false;

   // Initialisation
   sReqChamps_l := '';
   sAffChamps_l := '';
   sReqTables_l := StringReplace(sTable_p, ';', ',', [rfReplaceAll, rfIgnoreCase]);

   //Liste de tous les champs (visibles & invisibles)
   sLeChamp_l := Trim(ReadTokenSt(sChamps_p));
   sFormat_l := Trim(ReadTokenSt(sFormats_p));
   iOrdre_l := 0;

   while sLeChamp_l <> '' do
   begin
      iInd_l := 0;
      if sLeChamp_l[1] = '(' then
      begin
         Inc(iInd_l);
         sReqChamps_l := sReqChamps_l + sLeChamp_l + ' AS FORMULE' + IntToStr(iInd_l) + ', ';
         sAffChamps_l := sAffChamps_l + 'FORMULE' + IntToStr(iInd_l) + ';';
      end
      else
      begin
         sReqChamps_l := sReqChamps_l + sLeChamp_l + ', ';
         sAffChamps_l := sAffChamps_l + sleChamp_l + ';';
      end;

      if sFormat_l[5] = '$' then
      begin
         // Champ lié à une tablette dont on affiche le libellé complet
         if DonneesTablette(ChampToTT(sLeChamp_l), sCode_l, sChampLib_l,
                            sTable_l, sWhere_l, sPrefixe_l) then
         begin
            Inc(iOrdre_l);
            sOrdre_l := sPrefixe_l + IntToStr(iOrdre_l);

            sReqChamps_l := sReqChamps_l + sOrdre_l + '.' + sChampLib_l + ' AS ' + sOrdre_l + sChampLib_l + ', ';
            sAffChamps_l := sAffChamps_l + sOrdre_l + sChampLib_l + ';';

            sReqJoin_l := sReqJoin_l + ' LEFT JOIN ' + sTable_l + ' AS ' + sOrdre_l +
                                       ' ON ' + sLeChamp_l + ' = ' + sOrdre_l + '.' + sCode_l;

            sWhere_l := StringReplace(sWhere_l, sPrefixe_l + '_', sOrdre_l + '.' + sPrefixe_l + '_', [rfReplaceAll, rfIgnoreCase]);

            // Jointure avec la tablette
            if sWhere_l <> '' then
               sReqJoin_l := sReqJoin_l + ' AND ' + sWhere_l;
         end;

      end;
      sLeChamp_l := Trim(ReadTokenSt(sChamps_p));
      sFormat_l := Trim(ReadTokenSt(sFormats_p));
   end;

   if sReqChamps_l = '' then
      exit;

   // Requête
   sRequete_l := 'SELECT ' + Copy(sReqChamps_l, 1, Length(sReqChamps_l) - 2) + ' ' +
                 'FROM '   + sReqTables_l;

   if sReqJoin_l <> '' then
      sRequete_l := sRequete_l + sReqJoin_l;

   if sJointure_p <> '' then
   begin
      if sWhere_p <> '' then
         sRequete_l := sRequete_l + ' WHERE ' + sJointure_p + ' AND ' + sWhere_p
      else
         sRequete_l := sRequete_l + ' WHERE ' + sJointure_p;
   end
   else
   begin
      if sWhere_p <> '' then
         sRequete_l := sRequete_l + ' WHERE ' + sWhere_p;
   end;

   if sOrdre_p <> '' then
      sRequete_l := sRequete_l + ' ORDER BY ' + sOrdre_p;

   result := ChargeLaGrille(sRequete_l, sAffChamps_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 06/07/2004
Modifié le ... : 06/07/2004
Description .. : Chargement des données dans la grille depuis une requête
Mots clefs ... : 
*****************************************************************}
function  TGrilleSansMul.ChargeLaGrille(sRequete_p, sAffChamps_p : string) : boolean;
var
   iRowInd_l, iColInd_l : integer;
begin
   result := false;
   OBTable_c.ClearDetail;

   OBTable_c.LoadDetailFromSQL(sRequete_p, false, false);
   for iRowInd_l := 0 to OBTable_c.detail.count -1 do
      OBTable_c.Detail[iRowInd_l].AddChampSupValeur('IND', iRowInd_l);

   if sAffChamps_p = '' then
      sAffChamps_p := ExtractChamps(sRequete_p);

   gdGrille_c.VidePile(false);
   OBTable_c.PutGridDetail(gdGrille_c, true, true, 'IND;' + sAffChamps_p);
   PCumul_c.Caption := 'Totaux (' + IntToStr(NbRows) + ' lignes)';
   iColInd_l := 0;
   while READTOKENST(sAffChamps_p) <> '' do
      Inc(iColInd_l);
   if gdGrille_c.ColCount <> iColInd_l then
      gdGrille_c.ColCount := iColInd_l;   
   result := true;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 06/07/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function  TGrilleSansMul.ExtractChamps(sRequete_p : string) : string;
var
   sChamp_l, sLesChamps_l, sAffChamps_l : string;
begin
   sLesChamps_l := UpperCase(sRequete_p);
   READTOKENPipe(sLesChamps_l, 'SELECT ');
   sLesChamps_l := READTOKENPipe(sLesChamps_l, ' FROM');

   sChamp_l := Trim(READTOKENPipe(sLesChamps_l, ','));
   while sChamp_l <> '' do
   begin
      if pos(' AS ', sChamp_l) > 0 then
         READTOKENPipe(sChamp_l, ' AS ');

      if sAffChamps_l <> '' then
         sAffChamps_l := sAffChamps_l + ';';
      sAffChamps_l := sAffChamps_l + sChamp_l;

      sChamp_l := Trim(READTOKENPipe(sLesChamps_l, ','));
   end;
   result := sAffChamps_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 28/06/2004
Modifié le ... :   /  /
Description .. : Recherche de la tablette en fonction du suffixe
Mots clefs ... :
*****************************************************************}
function  TGrilleSansMul.DonneesTablette(sTablette_p : string;
                          var sCode_p, sChampLib_p, sTable_p, sWhere_p, sPrefixe_p : string) : boolean;
var
   OBTablette_l : TOB;
begin
   result := false;
   if sTablette_p = '' then
      Exit;
      
   OBTablette_l := TOB.Create('DECOMBOS', nil, -1);
   OBTablette_l.LoadDetailDBFromSQL('DECOMBOS', 'SELECT * FROM DECOMBOS ' +
                                    'WHERE DO_COMBO = "' + sTablette_p + '"');

   if OBTablette_l.Detail.Count > 0 then
   begin
      sTable_p    := PrefixeToTable(OBTablette_l.Detail[0].GetValue('DO_PREFIXE'));
      sCode_p     := OBTablette_l.Detail[0].GetValue('DO_CODE');
      sChampLib_p := OBTablette_l.Detail[0].GetValue('DO_CHAMPLIB');
      sWhere_p    := OBTablette_l.Detail[0].GetValue('DO_WHERE');
      sWhere_p    := StringReplace(sWhere_p, '&#@', '', [rfReplaceAll, rfIgnoreCase]);
      sPrefixe_p  := OBTablette_l.Detail[0].GetValue('DO_PREFIXE');
   end;
   OBTablette_l.Free;
   result := true;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 28/06/2004
Modifié le ... : 29/06/2004
Description .. : Formatage de l'affichage : largeur, titres,
Suite ........ : libelle tablette, ... des colonnes
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TGrilleSansMul.FormateLaGrille(sTitres_p, sLargeurs_p, sFormats_p : string);
var
   iColInd_l, iLargeur_l, iLargDefaut_l, iLargCols_l, iLargGrille_l : integer;
   sTitre_l, sFormat_l : string;
begin
   iLargGrille_l := gdGrille_c.Width;
   iLargDefaut_l := Round(iLargGrille_l / (gdGrille_c.ColCount * 10));
   if bAvecCpt_c then
   begin
      gdGrille_c.Cells[0, 0] := 'N°';
      gdGrille_c.ColWidths[0] := 2;
      iLargCols_l := 2;
   end
   else
   begin
      iLargCols_l := 0;
      gdGrille_c.ColWidths[0] := -1;
   end;

   iColInd_l := 1;
   While iColInd_l < gdGrille_c.ColCount do
   begin
      if sFormats_p = '' then
         sFormat_l := 'G.0O ---'
      else
         sFormat_l := Trim(ReadTokenSt(sFormats_p));

      if sLargeurs_p = '' then
         iLargeur_l := iLargDefaut_l
      else
         iLargeur_l := Round(ReadTokenI(sLargeurs_p));

      sTitre_l := Trim(ReadTokenSt(sTitres_p));

      if (iLargeur_l > 0) and (sFormat_l[6] = '-') then
      begin
         // Colonne visible
         if sFormat_l[5] = '$' then
         begin
            // Affichage du libellé complet de la tablette
            // Le champ code est invisible, on passe à la colonne suivante
            gdGrille_c.Cells[iColInd_l, 0] := gdGrille_c.ColNames[iColInd_l];
            gdGrille_c.ColWidths[iColInd_l] := -1;
            Inc(iColInd_l);
         end;

         iLargCols_l := iLargCols_l + iLargeur_l;
         gdGrille_c.ColWidths[iColInd_l] := iLargeur_l;
         
         if sTitre_l <> '' then
            gdGrille_c.Cells[iColInd_l, 0] := sTitre_l;

         if sFormat_l[1] = 'G' then
            gdGrille_c.ColAligns[iColInd_l] := taLeftJustify
         else if sFormat_l[1] = 'D' then
            gdGrille_c.ColAligns[iColInd_l] := taRightJustify
         else
            gdGrille_c.ColAligns[iColInd_l] := taCenter;
      end
      else
      begin
         // Colonne invisible
         gdGrille_c.ColWidths[iColInd_l] := -1;
      end;
      Inc(iColInd_l);
   end;

   // Ajustement des largeurs de colonnes en fonction de la largeur de la grille
   For iColInd_l := 0 to gdGrille_c.ColCount - 1 do
   begin
      if (gdGrille_c.ColWidths[iColInd_l] > -1) then
      begin
         iLargeur_l := Round(gdGrille_c.ColWidths[iColInd_l] * iLargGrille_l / iLargCols_l);
         gdGrille_c.ColWidths[iColInd_l] := iLargeur_l;
      end;
   end;

   gdGrille_c.Row := 1;
   iCurRow_c := 1;
end;

{procedure TGrilleSansMul.FormateLaGrille(sTitres_p, sLargeurs_p, sFormats_p : string);
var
   iColInd_l, iLargeur_l, iLargDefaut_l, iLargCols_l, iLargGrille_l : integer;
   sTitre_l, sFormat_l : string;
begin
   iLargGrille_l := gdGrille_c.Width;
   iLargDefaut_l := Round(iLargGrille_l / (gdGrille_c.ColCount * 10));
   if bAvecCpt_c then
   begin
      gdGrille_c.Cells[0, 0] := 'N°';
      gdGrille_c.ColWidths[0] := 2;
      iLargCols_l := 2;
   end
   else
   begin
      iLargCols_l := 0;
      gdGrille_c.ColWidths[0] := -1;
   end;

   iColInd_l := 1;
   While iColInd_l < gdGrille_c.ColCount do
   begin
      if sFormats_p = '' then
         sFormat_l := 'G.0O ---'
      else
         sFormat_l := Trim(ReadTokenSt(sFormats_p));

      if sLargeurs_p = '' then
         iLargeur_l := iLargDefaut_l
      else
         iLargeur_l := Round(ReadTokenI(sLargeurs_p));

      sTitre_l := Trim(ReadTokenSt(sTitres_p));

      if (iLargeur_l > 0) and (sFormat_l[6] = '-') then
      begin
         // Colonne visible
         if sFormat_l[5] = '$' then
         begin
            // Affichage du libellé complet de la tablette
            // Le champ code est invisible, on passe à la colonne suivante
            gdGrille_c.Cells[iColInd_l, 0] := gdGrille_c.ColNames[iColInd_l];
            gdGrille_c.ColWidths[iColInd_l] := -1;
            Inc(iColInd_l);
         end;

         iLargCols_l := iLargCols_l + iLargeur_l;
         gdGrille_c.ColWidths[iColInd_l] := iLargeur_l;
         
         if sTitre_l <> '' then
            gdGrille_c.Cells[iColInd_l, 0] := sTitre_l;

         if sFormat_l[1] = 'G' then
            gdGrille_c.ColAligns[iColInd_l] := taLeftJustify
         else if sFormat_l[1] = 'D' then
            gdGrille_c.ColAligns[iColInd_l] := taRightJustify
         else
            gdGrille_c.ColAligns[iColInd_l] := taCenter;
      end
      else
      begin
         // Colonne invisible
         gdGrille_c.ColWidths[iColInd_l] := -1;
      end;
      Inc(iColInd_l);
   end;

   // Ajustement des largeurs de colonnes en fonction de la largeur de la grille
   For iColInd_l := 0 to gdGrille_c.ColCount - 1 do
   begin
      if (gdGrille_c.ColWidths[iColInd_l] > -1) then
      begin
         iLargeur_l := Round(gdGrille_c.ColWidths[iColInd_l] * iLargGrille_l / iLargCols_l);
         gdGrille_c.ColWidths[iColInd_l] := iLargeur_l;
      end;
   end;
end;
}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 13/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGrilleSansMul.ListeCorrecte(sChamp_p, sLeTitre_p : String): Boolean;
begin
   Result := true;
//   if OBTable_c.GetNumChamp(sChamp_p) = -1 then
   If ChampEstDansListe(sChamp_p) = -1 then
   begin
      PGIInfo('La colonne "' + sLeTitre_p + '" ne figure pas dans votre paramètrage de liste.'+#13+#10 +
              'Veuillez la rajouter.', TitreHalley);
//      TButton(GetControl('BPARAMLISTE')).Click;
      result := false;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 13/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGrilleSansMul.ChampEstDansListe(sChamp_p : String): integer;
var
   sChamps_l, sChamp_l : string;
   iInd_l : integer;
begin
   Result := -1;
   sChamps_l := sChamps_c;
   iInd_l := 0;
   repeat
      sChamp_l := Trim(READTOKENST(sChamps_l));
      Inc(iInd_l);
   until (sChamp_l = '') or (sChamp_l = sChamp_p);

   if sChamp_l = sChamp_p then
      result := iInd_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 29/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGrilleSansMul.NbRows : integer;
begin
//   result := gdGrille_c.RowCount - gdGrille_c.FixedRows;
   result := OBTable_c.Detail.Count;
end;
/////////////////////////////////////////////////////////////////
end.

