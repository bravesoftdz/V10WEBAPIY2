{***********UNITE*************************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 23/06/2004
Modifié le ... :   /  /
Description .. : Chargement de données dans une grille non reliée à un MUL
Mots clefs ... :
*****************************************************************}
Unit dpJurOutilsGrille ;

Interface

Uses
   HEnt1, sysutils, Classes,

   {$IFDEF VER150}
   Variants,
   {$ENDIF}

   HCtrls, UTOB;

/////////////////////////////////////////////////////////////////
function  OnLoadGrille(gdGrille_p : ThGrid; sListe_p, sWhere_p : string;
                       bAvecCpt_p : boolean = false; iNbLignes_p : integer = -1) : integer; overload;

function  OnLoadGrille(gdGrille_p : ThGrid;
                       sRequete_p, sAffChamps_p, sTitres_p, slargeurs_p, sFormats_p : string;
                       bAvecCpt_p : boolean = false; iNbLignes_p : integer = -1) : integer; overload;


function  ChargeLaGrille(gdGrille_p : ThGrid;
                         sTable_p, sChamps_p, sWhere_p, sJointure_p,
                         sOrdre_p, sFormats_p : string;
                         bAvecCpt_p : boolean; iNbLignes_p : integer) : boolean; overload;

function  ChargeLaGrille(gdGrille_p : ThGrid;
                         sRequete_p, sAffChamps_p : string;
                         bAvecCpt_p : boolean; iNbLignes_p : integer) : boolean; overload;


function  ChargeLaListe(sListe_p : string; var tsMemo_p : TStrings) : boolean;

procedure FormateLaGrille(gdGrille_p : ThGrid;
                          sTitres_p, sLargeurs_p, sFormats_p : string;
                          bAvecCpt_p : boolean = false);

function  DonneesTablette(sTablette_p : string;
                          var sCode_p, sChampLib_p, sTable_p, sWhere_p, sPrefixe_p : string) : boolean;

function  ExtractChamps(sRequete_p : string) : string;

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
function OnLoadGrille(gdGrille_p : ThGrid; sListe_p, sWhere_p : string;
                       bAvecCpt_p : boolean = false; iNbLignes_p : integer = -1) : integer;
var
   tsMemo_l : TStrings;
   sTable_l, sChamps_l, sJointure_l, sOrdre_l, sTitres_l, sLargeurs_l, sFormats_l : string;
begin
   // Chargement du mémo
   tsMemo_l := TStringList.Create;
   if not ChargeLaListe(sListe_p, tsMemo_l) then
   begin
      tsMemo_l.Free;
      Exit;
   end;

   sTable_l    := UpperCase(Trim(tsMemo_l[0]));
   sChamps_l   := UpperCase(Trim(tsMemo_l[1]));
   sJointure_l := UpperCase(Trim(tsMemo_l[2]));
   sOrdre_l    := UpperCase(Trim(tsMemo_l[3]));
   sTitres_l   := Trim(tsMemo_l[4]);
   sLargeurs_l := UpperCase(Trim(tsMemo_l[5]));
   sFormats_l  := UpperCase(Trim(tsMemo_l[6]));

   if not ChargeLaGrille(gdGrille_p, sTable_l, sChamps_l, sWhere_p, sJointure_l,
                         sOrdre_l, sFormats_l, bAvecCpt_p, iNbLignes_p) then
   begin
      tsMemo_l.Free;
      Exit;
   end;

   FormateLaGrille(gdGrille_p, sTitres_l, sLargeurs_l, sFormats_l, bAvecCpt_p);

   tsMemo_l.free;
   result := gdGrille_p.RowCount - gdGrille_p.FixedRows;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 06/07/2004
Modifié le ... : 06/07/2004
Description .. : Chargement des données depuis une requête, affichage et
Suite ........ : formatage dans la grille
Mots clefs ... : 
*****************************************************************}
function OnLoadGrille(gdGrille_p : ThGrid;
                       sRequete_p, sAffChamps_p, sTitres_p, slargeurs_p, sFormats_p : string;
                       bAvecCpt_p : boolean = false; iNbLignes_p : integer = -1) : integer;
begin
   if ChargeLaGrille(gdGrille_p, sRequete_p, sAffChamps_p,
                     bAvecCpt_p, iNbLignes_p) then
      FormateLaGrille(gdGrille_p, sTitres_p, sLargeurs_p, sFormats_p, bAvecCpt_p);
   result := gdGrille_p.RowCount - gdGrille_p.FixedRows;      
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 28/06/2004
Modifié le ... :   /  /
Description .. : Chargement de la liste
Mots clefs ... :
*****************************************************************}
function  ChargeLaListe(sListe_p : string; var tsMemo_p : TStrings) : boolean;
var
   OBListe_l : TOB;
   sLesData_l, sData_l : string;
begin
   result := false;

   OBListe_l := TOB.Create('LISTE', nil, -1);
   OBListe_l.LoadDetailDBFromSQL('LISTE', 'SELECT LI_DATA,LI_UTILISATEUR FROM LISTE ' +
                                 'WHERE LI_LISTE = "' + sListe_p + '" ' +
                                 '  AND LI_UTILISATEUR = "' + V_PGI.User + '"');

   if OBListe_l.Detail.Count = 0 then
   begin
      OBListe_l.ClearDetail;
      OBListe_l.LoadDetailDBFromSQL('LISTE', 'SELECT LI_DATA,LI_UTILISATEUR FROM LISTE ' +
                                    'WHERE LI_LISTE = "' + sListe_p + '" ' +
                                    '  AND LI_UTILISATEUR = "---"');
      if OBListe_l.Detail.Count = 0 then
      begin
         OBListe_l.Free;
         Exit;
      end;
   end;

   // Chargement du mémo
   sLesData_l := VarAsType(OBListe_l.Detail[0].GetValue('LI_DATA'), varString);

   sData_l := READTOKENPipe(sLesData_l, #$D#$A);
   while sData_l <> '' do
   begin
      tsMemo_p.Add(sData_l);
      sData_l := READTOKENPipe(sLesData_l, #$D#$A)
   end;

   if tsMemo_p.Count < 9 then
   begin
      OBListe_l.Free;   //La liste est incohérente
      exit;
   end;
   OBListe_l.Free;
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
function  ChargeLaGrille(gdGrille_p : ThGrid;
                         sTable_p, sChamps_p, sWhere_p, sJointure_p,
                         sOrdre_p, sFormats_p : string;
                         bAvecCpt_p : boolean; iNbLignes_p : integer) : boolean;
var
   sReqChamps_l, sReqTables_l, sRequete_l, sReqJoin_l : string;
   sLeChamp_l, sAffChamps_l, sFormat_l, sPrefixe_l : string;
   sCode_l, sChampLib_l, sTable_l, sWhere_l, sOrdre_l : string;
   iOrdre_l : integer;
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
      if sLeChamp_l[1] = '(' then
      begin
         sLeChamp_l := Trim(ReadTokenSt(sChamps_p));
         sFormat_l := Trim(ReadTokenSt(sFormats_p));
         Continue;
      end;

      sReqChamps_l := sReqChamps_l + sLeChamp_l + ', ';
      sAffChamps_l := sAffChamps_l + sleChamp_l + ';';

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

   result := ChargeLaGrille(gdGrille_p, sRequete_l, sAffChamps_l,
                            bAvecCpt_p, iNbLignes_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 06/07/2004
Modifié le ... : 06/07/2004
Description .. : Chargement des données dans la grille depuis une requête
Mots clefs ... : 
*****************************************************************}
function  ChargeLaGrille(gdGrille_p : ThGrid;
                         sRequete_p, sAffChamps_p : string;
                         bAvecCpt_p : boolean; iNbLignes_p : integer) : boolean;
var
   OBTable_l : TOB;
   iOrdre_l : integer;
begin
   result := false;

   OBTable_l := TOB.Create('', nil, -1);
   OBTable_l.LoadDetailFromSQL(sRequete_p, false, false, iNbLignes_p);

   for iOrdre_l := 0 to OBTable_l.detail.count -1 do
      OBTable_l.detail[iOrdre_l].AddChampSupValeur('CPT', iOrdre_l + 1);

   if sAffChamps_p = '' then
      sAffChamps_p := ExtractChamps(sRequete_p);

   OBTable_l.PutGridDetail(gdGrille_p, true, true, 'CPT;' + sAffChamps_p);
   OBTable_l.Free;

   result := true;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 06/07/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function  ExtractChamps(sRequete_p : string) : string;
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
function  DonneesTablette(sTablette_p : string;
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
procedure FormateLaGrille(gdGrille_p : ThGrid;
                  sTitres_p, sLargeurs_p, sFormats_p : string;
                  bAvecCpt_p : boolean = false);
var
   iColInd_l, iLargeur_l, iLargDefaut_l, iLargCols_l, iLargGrille_l : integer;
   sTitre_l, sFormat_l : string;
begin
   iLargGrille_l := gdGrille_p.Width;
   iLargDefaut_l := Round(iLargGrille_l / (gdGrille_p.ColCount * 10));
   if bAvecCpt_p then
   begin
      gdGrille_p.Cells[0, 0] := 'N°';
      gdGrille_p.ColWidths[0] := 2;
      iLargCols_l := 2;
   end
   else
   begin
      iLargCols_l := 0;
      gdGrille_p.ColWidths[0] := -1;
   end;

   iColInd_l := 1;
   While iColInd_l < gdGrille_p.ColCount do
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
            gdGrille_p.Cells[iColInd_l, 0] := gdGrille_p.ColNames[iColInd_l];
            gdGrille_p.ColWidths[iColInd_l] := -1;
            Inc(iColInd_l);
         end;

         iLargCols_l := iLargCols_l + iLargeur_l;
         gdGrille_p.ColWidths[iColInd_l] := iLargeur_l;
         
         if sTitre_l <> '' then
            gdGrille_p.Cells[iColInd_l, 0] := sTitre_l;

         if sFormat_l[1] = 'G' then
            gdGrille_p.ColAligns[iColInd_l] := taLeftJustify
         else if sFormat_l[1] = 'D' then
            gdGrille_p.ColAligns[iColInd_l] := taRightJustify
         else
            gdGrille_p.ColAligns[iColInd_l] := taCenter;
      end
      else
      begin
         // Colonne invisible
         gdGrille_p.ColWidths[iColInd_l] := -1;
      end;
      Inc(iColInd_l);
   end;

   // Ajustement des largeurs de colonnes en fonction de la largeur de la grille
   For iColInd_l := 0 to gdGrille_p.ColCount - 1 do
   begin
      if (gdGrille_p.ColWidths[iColInd_l] > -1) then
      begin
         iLargeur_l := Round(gdGrille_p.ColWidths[iColInd_l] * iLargGrille_l / iLargCols_l);
         gdGrille_p.ColWidths[iColInd_l] := iLargeur_l;
      end;
   end;
end;
{function TOF_ANNULIEN.ListeCorrecte(nomchp, titrechp: String): Boolean;
begin
  Result := True;
  //  if Not GetDataSet.FieldExists(nomchp) then => pas possible (du moins si pas d'enreg...)
  If Not ChampEstDansQuery(nomchp, TFMul(Ecran).Q) then
    begin
    Result := False;
    PGIInfo('La colonne "'+titrechp+'" ne figure pas dans votre paramètrage de liste.'+#13+#10
     +'Veuillez la rajouter dans "Afficher les colonnes suivantes".', TitreHalley);
    TButton(GetControl('BPARAMLISTE')).Click;
    end;
end;}

/////////////////////////////////////////////////////////////////
end.

