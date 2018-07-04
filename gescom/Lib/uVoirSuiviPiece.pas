unit UVOIRSUIVIPIECE;

interface

uses
  uTob,
  EntGC, FactTOB,uEntCommun,UtilTOBPiece {pour le type R_CleDoc}
  ;

function Charge_Tob_Commande(nComm: Integer; CDRef: R_CleDoc; iIndice: Integer; sWhere: string = ''): tob;
function Charge_Tob_Client(cClie: string; TypeSortie: string = 'E'; sWhere: string = ''): tob;
function Charge_Tob_Fournisseur(cFour: string; TypeSortie: string = 'E'; sWhere: string = ''): tob;

procedure Charge_Arborescence_Pieces(TobLignes: Tob; sTypeSortie: string = 'E');
function Mets_A_Plat(TobLignes: Tob; lAchat: Boolean): Tob;
{V500_006}
function MakeReference(TobLigne: Tob; lFull: Boolean = True; sTypeSortie: string = 'E'; lOrdre: Boolean = False): string;

procedure AglPiecesClient(Parms: array of variant; Nb: Integer);
procedure AglPiecesFournisseur(Parms: array of variant; Nb: Integer);
function Genere_Tob(nPiece: Integer; CDRef: R_CleDoc; TypeSortie: string; sWhere: string = ''): Tob;
function Genere_Tob_Client(sClient: string; TypeSortie: string = 'E'; sWhere: string = ''): Tob;
function Genere_Tob_Fournisseur(sFournisseur: string; TypeSortie: string = 'E'; sWhere: string = ''): Tob;
procedure VoirPiece(nPiece: Integer; CDRef: R_CleDoc; TypeSortie: string);
procedure RegisterProc;



implementation

uses {Wcommuns,}
  Sysutils,
  Dialogs,
  M3fp,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  edtRetat,
  FE_MAIN, // Pour AGLLANCEFICHE
  {$ELSE}
  utilEagl,
  MainEagl, // Pour AglLanceGiche en eAgl
  {$ENDIF}
  hctRls,
  FactUtil,
  FactComm,
  hMsgBox,
  hEnt1,
  Wcommuns;
  //wTobDebug_Tof;

// Liste des champs chargés par les requettes dans la TOB. Seuls ces champs peuvent être affichés sur les voirs.
// Si on peut sans problèmes en rajouter, il fait faire attention à ne pas supprimer n'importe quels champs :
// certains, non visibles, sont utilisés dans le programme.
// Ne pas supprimer les champs suivants : GL_PIECEPRECEDENTE,GL_SOUCHE,GL_INDICEG,GL_NUMLIGNE,GL_NATUREPIECEG,GL_NUMERO,GL_DATEPIECE, GL_NUMORDRE
// ne pas supprimer non plus l'espace de fin.

const sListFields =
 'GL_PIECEPRECEDENTE,GL_SOUCHE,GL_INDICEG,GL_NUMLIGNE,GL_NUMORDRE,GL_NATUREPIECEG,GL_NUMERO,GL_DATEPIECE,GL_CODEARTICLE,GL_QTESTOCK,GL_QTERESTE,GL_QUALIFQTESTO,GL_PUHTDEV,GL_DEVISE,GL_MONTANTHTDEV,GL_DATELIVRAISON,GL_TIERS, GL_ARTICLE, GL_LIBELLE ';
  iDecal = 3;


var sListNPClient, sListNPFournisseur : string;


procedure MakeListes;
var i, PosVenteAchat, PosNaturePieceg : Integer;
begin
  sListNPClient := '';
  sListNPFournisseur := '';
  PosVenteAchat := VH_GC.TobParPiece.Detail[0].GetNumChamp('GPP_VENTEACHAT');
  PosNaturePieceg := VH_GC.TobParPiece.Detail[0].GetNumChamp('GPP_NATUREPIECEG');
  for i := 0 to VH_GC.TobParPiece.detail.Count - 1 do
  begin
    if VH_GC.TobParPiece.detail[i].Getvaleur(PosVenteAchat) = 'VEN' then
      sListNPClient := sListNPClient+'"'+ VH_GC.TobParPiece.detail[i].Getvaleur(PosNaturePieceg)+'",'
    else if VH_GC.TobParPiece.detail[i].Getvaleur(PosVenteAchat) = 'ACH' then
      sListNPFournisseur := sListNPFournisseur+'"'+ VH_GC.TobParPiece.detail[i].Getvaleur(PosNaturePieceg)+'",';
  end;
  sListNPClient := Copy(sListNPClient,1,Length(sListNPClient)-1);
  sListNPFournisseur := Copy(sListNPFournisseur,1,Length(sListNPFournisseur)-1);
end;

function LoadTobFromSql(TableName, Sql: string; t: Tob; lAppend: boolean = false; NbLigne: integer = -1): boolean;
var
  Q: tQuery;
begin
  Result := false;
  Q := OpenSql(Sql, True, NbLigne);
  try
    if not Q.Eof then
    begin
      t.loadDetailDB(TableName, '', '', Q, lAppend);
      Result := true;
    end;
  finally
    Ferme(Q);
  end;
end;

{Fonctions génériques de manipulation de chaines}

function Replicate(s: string; iLong: integer): string;
var
  i: integer;
begin
  Result := '';
  for i := 1 to iLong do
    Result := Result + s;
end;

function PadLeft(sChaine: string; iLong: integer; sCaractere: string = ' '): string;
begin
  if Length(sChaine) < iLong then
    Result := Replicate(sCaractere, iLong - Length(sChaine)) + sChaine
  else
    Result := Copy(sChaine, 1, iLong);
end;

{Fin Fonctions génériques de manipulation de chaines}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Pascal Gautherot
Créé le ...... : 17/12/2002
Modifié le ... :   /  /
Description .. : Fonction de chargement dans une TOB des lignes d'une
Suite ........ : commande de vente.
Suite ........ : Recoit en paramètre le N° de commande
Suite ........ : Retourne une TOB contenant toutes les lignes "de première
Suite ........ : génération" (non issues d'une transformation de pièces)
Mots clefs ... : LIGNE, TOB
*****************************************************************}

function Charge_Tob_Commande(nComm: Integer; CDRef: R_CleDoc; iIndice: Integer; sWhere: string = ''): tob;
var TobLignes: Tob;
  sQuery, sWhereCle: string;

begin
  TobLignes := Tob.Create('LIGNE', nil, -1);
  sWhereCle := WherePiece(CDRef, ttdLigne, False);
  sQuery := 'SELECT ' + sListFields +
            ', GA.GA_LIBELLE, D.D_SYMBOLE, DI1.GDI_LIBELLE AS GDI_DIM1, DI2.GDI_LIBELLE AS GDI_DIM2, DI3.GDI_LIBELLE AS GDI_DIM3,DI4.GDI_LIBELLE AS GDI_DIM4, DI5.GDI_LIBELLE AS GDI_DIM5'+
            ' From LIGNE'+
            ' left join ARTICLE GA on GA.GA_ARTICLE = GL_ARTICLE' +
            ' left join DEVISE D on D.D_DEVISE = GL_DEVISE' +
            ' left join dimension di1 on di1.gdi_typedim ="DI1" and di1.gdi_grilledim =GA.ga_grilledim1 and di1.gdi_codedim =GA.ga_codedim1'+
            ' left join dimension di2 on di2.gdi_typedim ="DI2" and di2.gdi_grilledim =GA.ga_grilledim2 and di2.gdi_codedim =GA.ga_codedim2'+
            ' left join dimension di3 on di3.gdi_typedim ="DI3" and di3.gdi_grilledim =GA.ga_grilledim3 and di3.gdi_codedim =GA.ga_codedim3'+
            ' left join dimension di4 on di4.gdi_typedim ="DI4" and di4.gdi_grilledim =GA.ga_grilledim4 and di4.gdi_codedim =GA.ga_codedim4'+
            ' left join dimension di5 on di5.gdi_typedim ="DI5" and di5.gdi_grilledim =GA.ga_grilledim5 and di5.gdi_codedim =GA.ga_codedim5'+
            ' WHERE ' + sWhereCle + ' and GL_QTESTOCK <> 0 ' + sWhere +' and GL_TYPELIGNE = "ART"';  
  LoadTobFromSql('LIGNE', sQuery, TobLignes);
  Result := TobLignes;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Pascal Gautherot
Créé le ...... : 17/12/2002
Modifié le ... : 06/01/2003
Description .. : Fonction de chargement dans une TOB des lignes d'un
Suite ........ : client.
Suite ........ : Recoit en paramètre le code client
Suite ........ : Retourne une TOB contenant toutes les lignes "de première
Suite ........ : génération" (non issues d'une transformation de pièces)
Mots clefs ... : LIGNE, TOB
*****************************************************************}

function Charge_Tob_Client(cClie: string; TypeSortie: string = 'E'; sWhere: string = ''): tob;
var TobLignes: Tob;
  srequ : string;
begin
  TobLignes := Tob.Create('LIGNE', nil, -1);
  srequ := 'SELECT ' + sListFields + ', GA.GA_LIBELLE, D.D_SYMBOLE, DI1.GDI_LIBELLE AS GDI_DIM1, DI2.GDI_LIBELLE AS GDI_DIM2, DI3.GDI_LIBELLE AS GDI_DIM3,DI4.GDI_LIBELLE AS GDI_DIM4, DI5.GDI_LIBELLE AS GDI_DIM5'+
            ' From LIGNE'+
            ' left join ARTICLE GA on GA.GA_ARTICLE = GL_ARTICLE' +
            ' left join DEVISE D on D.D_DEVISE = GL_DEVISE' +
            ' left join dimension di1 on di1.gdi_typedim ="DI1" and di1.gdi_grilledim =GA.ga_grilledim1 and di1.gdi_codedim =GA.ga_codedim1'+
            ' left join dimension di2 on di2.gdi_typedim ="DI2" and di2.gdi_grilledim =GA.ga_grilledim2 and di2.gdi_codedim =GA.ga_codedim2'+
            ' left join dimension di3 on di3.gdi_typedim ="DI3" and di3.gdi_grilledim =GA.ga_grilledim3 and di3.gdi_codedim =GA.ga_codedim3'+
            ' left join dimension di4 on di4.gdi_typedim ="DI4" and di4.gdi_grilledim =GA.ga_grilledim4 and di4.gdi_codedim =GA.ga_codedim4'+
            ' left join dimension di5 on di5.gdi_typedim ="DI5" and di5.gdi_grilledim =GA.ga_grilledim5 and di5.gdi_codedim =GA.ga_codedim5'+
            ' WHERE GL_QTESTOCK <> 0 AND (GL_NATUREPIECEG in (' + sListNPClient +
    '))and GL_PIECEPRECEDENTE = "" and GL_TIERS = "' + cClie + '"' + sWhere+' and GL_TYPELIGNE = "ART"';
  if LoadTobFromSql('LIGNE',srequ , TobLignes) then
    Result := TobLignes
  else
    Result := nil;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Pascal Gautherot
Créé le ...... : 15/01/2003
Modifié le ... :   /  /
Description .. : Chargement dans une tob des lignes "de premier niveau"
Suite ........ : (non issues de la transformation d'une pièce) d'un
Suite ........ : fournisseur
Suite ........ : Recoit en paramètre le code fournisseur
Suite ........ : Retourne une Tob contenant les lignes désirées.
Mots clefs ... :
*****************************************************************}

function Charge_Tob_Fournisseur(cFour: string; TypeSortie: string = 'E'; sWhere: string = ''): tob;
var TobLignes: Tob;
  srequ : String;
begin
  TobLignes := Tob.Create('LIGNE', nil, -1);
  srequ := 'SELECT ' + sListFields + ', GA.GA_LIBELLE, D.D_SYMBOLE, DI1.GDI_LIBELLE AS GDI_DIM1, DI2.GDI_LIBELLE AS GDI_DIM2, DI3.GDI_LIBELLE AS GDI_DIM3,DI4.GDI_LIBELLE AS GDI_DIM4, DI5.GDI_LIBELLE AS GDI_DIM5'+
            ' From LIGNE'+
            ' left join ARTICLE GA on GA.GA_ARTICLE = GL_ARTICLE' +
            ' left join DEVISE D on D.D_DEVISE = GL_DEVISE' +
            ' left join dimension di1 on di1.gdi_typedim ="DI1" and di1.gdi_grilledim =GA.ga_grilledim1 and di1.gdi_codedim =GA.ga_codedim1'+
            ' left join dimension di2 on di2.gdi_typedim ="DI2" and di2.gdi_grilledim =GA.ga_grilledim2 and di2.gdi_codedim =GA.ga_codedim2'+
            ' left join dimension di3 on di3.gdi_typedim ="DI3" and di3.gdi_grilledim =GA.ga_grilledim3 and di3.gdi_codedim =GA.ga_codedim3'+
            ' left join dimension di4 on di4.gdi_typedim ="DI4" and di4.gdi_grilledim =GA.ga_grilledim4 and di4.gdi_codedim =GA.ga_codedim4'+
            ' left join dimension di5 on di5.gdi_typedim ="DI5" and di5.gdi_grilledim =GA.ga_grilledim5 and di5.gdi_codedim =GA.ga_codedim5'+
            ' WHERE GL_QTESTOCK <> 0 AND (GL_NATUREPIECEG IN (' + sListNPFournisseur +
            '))and GL_PIECEPRECEDENTE = "" and GL_TIERS = "' + cFour + '"' + sWhere +' and GL_TYPELIGNE = "ART"';
  if LoadTobFromSql('LIGNE',srequ , TobLignes) then
    Result := TobLignes
  else
    Result := nil;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Pascal Gautherot
Créé le ...... : 19/12/2002
Modifié le ... :   /  /
Description .. : Retourne la valeur qui serait mémorisée dans
Suite ........ : GL_PIECEPRECEDENTE des lignes gérérées si cette
Suite ........ : pièce était tansformée.
Mots clefs ... :
*****************************************************************}
{V500_006 Début}
function MakeReference(TobLigne: Tob; lFull: Boolean = True; sTypeSortie: string = 'E'; lOrdre: Boolean = False): string;
var sNumLigne: string;
  sTmp, sTmp1 : string;
  I : Integer;
begin
  Result := '';
  if lFull then
  begin
    if lOrdre then
    begin
      sTmp := EncodeRefPiece(TobLigne);
      //PG : pour des raisons d'ordre de tri, on est obligé de cadrer les N° de pièce, indice et N° de ligne à droite.
      For i := 1 to 6 do
      Begin
        sTmp1 := ReadTokenSt(sTmp);
        If (i = 4) or (i = 5) or (i = 6) then
        begin
          sTmp1 := PadLeft(sTmp1,15,'0');
        end;
        Result := result + sTmp1+';';
      end;
    end
    else
    Result := EncodeRefPiece(TobLigne);
  end
  else
  begin
    if sTypeSortie = 'E' then
      sNumLigne := ''
    else
      sNumLigne := PadLeft(string(TobLigne.getValue('GL_NUMORDRE')),15,'0') + ';';
    // Pour avoir un tri chronologique, la date est formatée Année/mois/jours
    Result := FormatDateTime('yyyymmdd',TobLigne.GetValue('GL_DATEPIECE')) +
      string(TobLigne.GetValue('GL_NATUREPIECEG')) + ';' +
      string(TobLigne.GetValue('GL_SOUCHE')) + ';' +
      PadLeft(string(TobLigne.getValue('GL_NUMERO')),15,'0') + ';' +
      sNumLigne;
  end;
end;
{V500_006 Fin}

function MakeListeNatures(NaturesSuivantes: string): string;
var Liste: string;
begin
  Liste := '"' + NaturesSuivantes;
  if Copy(Liste, Length(Liste), 1) = ';' then
    Liste := Copy(Liste, 1, Length(Liste) - 1);
  Liste := '(' + StringReplace(Liste, ';', '","', [rfReplaceAll]) + '")';
  Result := Liste;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Pascal Gautherot
Créé le ...... : 17/12/2002
Modifié le ... :   /  /
Description .. : Procedure récursive de chargement dans une TOB de
Suite ........ : l'arborescence des sous pièces issues des pièces
Suite ........ : contenues dans la TOB recue en paramètres.
Mots clefs ... : TOB, LIGNE
*****************************************************************}

procedure Charge_Arborescence_Pieces(TobLignes: Tob; sTypeSortie: string = 'E');
var i, iNiv: integer;
  sRefPiece: string;
  sPieceOrig: string;
  sShortPieceOrig: string;
  iOrdre: Integer;
  idecalage: Integer;
  sNaturesSuivantes: string;

  procedure Cree_Arborescence(TobLignes: Tob);
  var sRefPiece: string;
    I: Integer;
    lInc: Boolean;
    sNaturesSuivantes: string;
  begin
    lInc := False;
    if pos(Toblignes.GetValue('GL_NATUREPIECEG'), TobLignes.GetValue('GL_PIECEPRECEDENTE')) = 0 then
    begin
      lInc := True;
      iNiv := iNiv + 1; // Niveau dans l'arborescence, en faisant abstraction des lignes de reliquat.  Utilisé dans les états.
    end;
    for i := 0 to TobLignes.Detail.Count - 1 do
    begin
      // 1) On crée un champs contenant l'équivalent du GL_PIECEPRECEDENTE des lignes que je veux charger
      sRefPiece := Makereference(TobLignes.detail[I], True, sTypeSortie);
      // 2) On récupère la liste des natures de pièces suivantes possibles, pour optimiser la requète.
      sNaturesSuivantes := MakeListeNatures(GetInfoParPiece(TobLignes.Detail[I].getValue('GL_NATUREPIECEG'), 'GPP_NATURESUIVANTE'));
      // 3) On charge les données dans la sous-tob
{      LoadTobFromSql('LIGNE', 'SELECT ' + sListFields + ' FROM LIGNE WHERE GL_ARTICLE="' + TobLignes.Detail[I].GetValue('GL_ARTICLE') + '" AND GL_TIERS="' +
        TobLignes.Detail[I].GetValue('GL_TIERS') + '" AND GL_NATUREPIECEG IN ' + sNaturesSuivantes + ' AND GL_QTESTOCK <> 0 AND GL_PIECEPRECEDENTE = "' +
        sRefPiece + '"', TobLignes.detail[I]);}
      LoadTobFromSql('LIGNE', 'SELECT ' + sListFields +
                  ', GA.GA_LIBELLE, D.D_SYMBOLE, DI1.GDI_LIBELLE AS GDI_DIM1, DI2.GDI_LIBELLE AS GDI_DIM2, DI3.GDI_LIBELLE AS GDI_DIM3,DI4.GDI_LIBELLE AS GDI_DIM4, DI5.GDI_LIBELLE AS GDI_DIM5'+
                  ' From LIGNE'+
                  ' left join ARTICLE GA on GA.GA_ARTICLE = GL_ARTICLE' +
                  ' left join DEVISE D on D.D_DEVISE = GL_DEVISE' +
                  ' left join dimension di1 on di1.gdi_typedim ="DI1" and di1.gdi_grilledim =GA.ga_grilledim1 and di1.gdi_codedim =GA.ga_codedim1'+
                  ' left join dimension di2 on di2.gdi_typedim ="DI2" and di2.gdi_grilledim =GA.ga_grilledim2 and di2.gdi_codedim =GA.ga_codedim2'+
                  ' left join dimension di3 on di3.gdi_typedim ="DI3" and di3.gdi_grilledim =GA.ga_grilledim3 and di3.gdi_codedim =GA.ga_codedim3'+
                  ' left join dimension di4 on di4.gdi_typedim ="DI4" and di4.gdi_grilledim =GA.ga_grilledim4 and di4.gdi_codedim =GA.ga_codedim4'+
                  ' left join dimension di5 on di5.gdi_typedim ="DI5" and di5.gdi_grilledim =GA.ga_grilledim5 and di5.gdi_codedim =GA.ga_codedim5'+
                  ' WHERE GL_ARTICLE="' + TobLignes.Detail[I].GetValue('GL_ARTICLE') +
        '" AND GL_NATUREPIECEG IN ' + sNaturesSuivantes + ' AND GL_QTESTOCK <> 0 AND GL_PIECEPRECEDENTE = "' +
        sRefPiece + '"', TobLignes.detail[I]);
      TobLignes.Detail[I].AddChampSupValeur('_ORIG', sPieceOrig);
      TobLignes.Detail[I].AddChampSupValeur('_ORIGCDE', sShortPieceOrig);
      TobLignes.Detail[I].AddChampSupValeur('_LIBELLE', Replicate(' ', (iNiv - 1) * idecalage) +
        GetInfoParPiece(TobLignes.Detail[I].GetValue('GL_NATUREPIECEG'), 'GPP_LIBELLE'));
      iOrdre := iOrdre + 1;
      TobLignes.Detail[I].AddChampSupValeur('_IORDRE', PadLeft(IntToStr(iOrdre), 4));
       TobLignes.Detail[I].AddChampSupValeur('_INIV', iNIV);
      // Puis on rappel la fonction de manière récursive pour aller chercher les sous niveaux.
      Cree_Arborescence(TobLignes.detail[I])
    end;
    if lInc then iNiv := iNiv - 1;
  end;

begin
  iOrdre := 0;
  iNiv := 1;
  if sTypeSortie = 'TV' then iDecalage := iDecal else iDecalage := 1;
  for i := 0 to TobLignes.Detail.Count - 1 do
  begin
    //  On crée un champs contenant l'équivalent du GL_PIECEPRECEDENTE des lignes que je veux charger
    sRefPiece := MakeReference(TobLignes.detail[I], True, sTypeSortie);
{V500_006}
    sPieceOrig := MakeReference(TobLignes.detail[I], True, sTypeSortie,True);
    sShortPieceOrig := MakeReference(TobLignes.detail[I], False, sTypeSortie);
    sNaturesSuivantes := MakeListeNatures(GetInfoParPiece(TobLignes.Detail[I].getValue('GL_NATUREPIECEG'), 'GPP_NATURESUIVANTE'));
    // On charge les données dans la sous-tob
{    LoadTobFromSql('LIGNE', 'SELECT ' + sListFields + ' FROM LIGNE WHERE GL_ARTICLE="' + TobLignes.Detail[I].GetValue('GL_ARTICLE') + '" AND GL_TIERS="' +
      TobLignes.Detail[I].GetValue('GL_TIERS') + '" AND GL_NATUREPIECEG IN ' + sNaturesSuivantes + ' AND GL_QTESTOCK <> 0 AND GL_PIECEPRECEDENTE = "' + sRefPiece
      + '"', TobLignes.detail[I]); }
    LoadTobFromSql('LIGNE', 'SELECT ' + sListFields +
            ', GA.GA_LIBELLE, D.D_SYMBOLE, DI1.GDI_LIBELLE AS GDI_DIM1, DI2.GDI_LIBELLE AS GDI_DIM2, DI3.GDI_LIBELLE AS GDI_DIM3,DI4.GDI_LIBELLE AS GDI_DIM4, DI5.GDI_LIBELLE AS GDI_DIM5'+
            ' From LIGNE'+
            ' left join ARTICLE GA on GA.GA_ARTICLE = GL_ARTICLE' +
            ' left join DEVISE D on D.D_DEVISE = GL_DEVISE' +
            ' left join dimension di1 on di1.gdi_typedim ="DI1" and di1.gdi_grilledim =GA.ga_grilledim1 and di1.gdi_codedim =GA.ga_codedim1'+
            ' left join dimension di2 on di2.gdi_typedim ="DI2" and di2.gdi_grilledim =GA.ga_grilledim2 and di2.gdi_codedim =GA.ga_codedim2'+
            ' left join dimension di3 on di3.gdi_typedim ="DI3" and di3.gdi_grilledim =GA.ga_grilledim3 and di3.gdi_codedim =GA.ga_codedim3'+
            ' left join dimension di4 on di4.gdi_typedim ="DI4" and di4.gdi_grilledim =GA.ga_grilledim4 and di4.gdi_codedim =GA.ga_codedim4'+
            ' left join dimension di5 on di5.gdi_typedim ="DI5" and di5.gdi_grilledim =GA.ga_grilledim5 and di5.gdi_codedim =GA.ga_codedim5'+
            ' WHERE GL_ARTICLE="' + TobLignes.Detail[I].GetValue('GL_ARTICLE') +
            '" AND GL_NATUREPIECEG IN ' + sNaturesSuivantes + ' AND GL_QTESTOCK <> 0 AND GL_PIECEPRECEDENTE = "' + sRefPiece +
            '"', TobLignes.detail[I]);
     TobLignes.Detail[I].AddChampSupValeur('_ORIG', sPieceOrig);
    TobLignes.Detail[I].AddChampSupValeur('_ORIGCDE', sShortPieceOrig);
    TobLignes.Detail[I].AddChampSupValeur('_LIBELLE', GetInfoParPiece(TobLignes.Detail[I].GetValue('GL_NATUREPIECEG'), 'GPP_LIBELLE'));
    iOrdre := iOrdre + 1;
    TobLignes.Detail[I].AddChampSupValeur('_IORDRE', PadLeft(IntToStr(iOrdre), 4));
    TobLignes.Detail[I].AddChampSupValeur('_INIV', iniv);
    Cree_Arborescence(TobLignes.detail[I])
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Pascal Gautherot
Créé le ...... : 18/12/2002
Modifié le ... :   /  /
Description .. : Mise à plat sur un seul niveau d'une tob hierarchique
Mots clefs ... : TOB
*****************************************************************}

function Mets_A_Plat(TobLignes: Tob; lAchat: Boolean): Tob;
var TobTmp: Tob;
  OrigCde: string;
  iRupture, i: Integer;
  sRupture: string;

  // Procédure récursive de mise à plat de la tob hierarchique.
  procedure DecomposeTob(MaTob: Tob);
  var I: Integer;
    T: Tob;
  begin
    for i := 0 to MatOb.detail.Count - 1 do
    begin
      T := Tob.create('LIGNE', TOBTMP, -1);
      TobTmp.AddChampSup('GA_LIBELLE',false);
      TobTmp.AddChampSup('D_SYMBOLE',false);
      TobTmp.AddChampSup('GDI_DIM1',false);
      TobTmp.AddChampSup('GDI_DIM2',false);
      TobTmp.AddChampSup('GDI_DIM3',false);
      TobTmp.AddChampSup('GDI_DIM4',false);
      TobTmp.AddChampSup('GDI_DIM5',false);
      TobTmp.AddChampSup('_ORIG', False);
      TobTmp.AddChampSup('_ORIGCDE', False);
      TobTmp.AddChampSup('_IORDRE', False);
      TobTmp.AddChampSup('_INIV', False);
      TobTmp.AddCHampSup('_LIBELLE', False);
      T.Dupliquer(MaTob.detail[I], False, True);
      if MaTob.detail[I].Detail.Count > 0 then
      begin
        DecomposeTob(MaTob.detail[I]);
      end;
    end;
  end;

begin
  TobTmp := Tob.Create('LIGNE', nil, -1);
  TobTmp.AddChampSup('GA_LIBELLE',false);
  TobTmp.AddChampSup('D_SYMBOLE',false);
  TobTmp.AddChampSup('GDI_DIM1',false);
  TobTmp.AddChampSup('GDI_DIM2',false);
  TobTmp.AddChampSup('GDI_DIM3',false);
  TobTmp.AddChampSup('GDI_DIM4',false);
  TobTmp.AddChampSup('GDI_DIM5',false);
  TobTmp.AddChampSup('_ORIG', False);
  TobTmp.AddChampSup('_ORIGCDE', False);
  TobTmp.AddChampSup('_IORDRE', False);
  TobTmp.AddChampSup('_INIV', False);
  TobTmp.AddCHampSup('_LIBELLE', False);
  TobTmp.AddCHampSup('_RUPTURE', False);
  DecomposeTob(TobLignes); // Fait la mise à plat

  if TobTmp.Detail.Count > 0 then
  begin
    TobTmp.Detail.Sort('_ORIGCDE;_ORIG;_IORDRE');
    OrigCde := TobTmp.Detail[0].GetValue('_ORIGCDE');
    iRupture := 1;
    I := 0;
    sRupture := '(' + WpadLeft(IntToStr(iRupture),8,' ') + ')   ' + TobTmp.Detail[I].GetValue('_LIBELLE') + ' ' + string(TobTmp.Detail[I].GetValue('GL_NUMERO')) + ' Ligne ' +
      string(TobTmp.Detail[I].GetValue('GL_NUMORDRE')) + ' Tiers : ' + string(TobTmp.Detail[I].GetValue('GL_TIERS'))+' '+rechdom('GCTIERS',TobTmp.Detail[I].GetValue('GL_TIERS'),False);
    for i := 0 to TobTmp.detail.Count - 1 do
    begin
      if OrigCde <> TobTmp.Detail[i].GetValue('_ORIGCDE') then
      begin
        OrigCde := TobTmp.Detail[i].GetValue('_ORIGCDE');
        iRupture := iRupture + 1;
        sRupture := '(' + WpadLeft(IntToStr(iRupture),8,' ') + ')   ' + TobTmp.Detail[I].GetValue('_LIBELLE') + ' ' + string(TobTmp.Detail[I].GetValue('GL_NUMERO')) + ' Ligne '
          + string(TobTmp.Detail[I].GetValue('GL_NUMORDRE')) + ' Tiers : ' + string(TobTmp.Detail[I].GetValue('GL_TIERS'))+' '+rechdom('GCTIERS',TobTmp.Detail[I].GetValue('GL_TIERS'),False);
      end;
      TobTmp.Detail[I].AddChampSupValeur('_RUPTURE', sRupture);
    end;
  end;
  Result := TobTmp;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Pascal Gautherot
Créé le ...... : 15/01/2003
Modifié le ... :   /  /
Description .. : Visualisation synthétique des pièces d'un client.
Suite ........ : Appel depuis le script de la fiche GCTIERS
Mots clefs ... :
*****************************************************************}

procedure AglPiecesClient(Parms: array of variant; Nb: Integer);
begin
  AglLanceFiche('GC', 'GCSUIVIPIECE', '', '', ';;;;;TV;' + string(Parms[0]) + ';CL') // Passer les elements de la clé en paramètre
end;

{***********A.G.L.***********************************************
Auteur  ...... : Pascal Gautherot
Créé le ...... : 15/01/2003
Modifié le ... : 15/01/2003
Description .. : Visualisation synthétique des pièces d'un fournisseur.
Suite ........ : Appel depuis le script de la fiche GCFOURNISSEUR
Mots clefs ... :
*****************************************************************}

procedure AglPiecesFournisseur(Parms: array of variant; Nb: Integer);
begin
  AglLanceFiche('GC', 'GCSUIVIPIECE', '', '', ';;;;;TV;' + string(Parms[0]) + ';FO') // Passer les elements de la clé en paramètre
end;

{***********A.G.L.***********************************************
Auteur  ...... : Pascal Gautherot
Créé le ...... : 15/01/2003
Modifié le ... :   /  /
Description .. : Visualisation synthétique des lignes d'une piece et des
Suite ........ : lignes générées par la transformation de celles ci
Mots clefs ... :
*****************************************************************}

procedure VoirPiece(nPiece: Integer; CDRef: R_CleDoc; TypeSortie: string);
begin
  AglLanceFiche('GC', 'GCSUIVIPIECE', '', '', IntToSTr(nPiece) + ';' + CDRef.NaturePiece + ';' + CDRef.Souche + ';' + IntToSTr(CDRef.NumeroPiece) + ';' +
    IntToSTr(CDRef.Indice) + ';' + TypeSortie) // Passer les elements de la clé en paramètre
end;

function Genere_Tob(nPiece: Integer; CDRef: R_CleDoc; TypeSortie: string; sWhere: string = ''): Tob;
var
  Toblignes, TobAplat: Tob;
  lAchat: Boolean;
begin
  TobAplat := nil;
  MakeListes;
  if CDRef.Indice = 0 then
  begin
    TobLignes := Charge_Tob_Commande(nPiece, CDref, CDRef.Indice, sWhere);
    if Assigned(TobLignes) then
    begin
      Charge_Arborescence_Pieces(TobLignes, TypeSortie);
      lAchat := (pos(CDref.NaturePiece, 'CC_BLC_FAC') = 0);
      TobAplat := Mets_A_Plat(TobLignes, lAchat);
    end;
  end
  else
    HShowMessage('0;'+TraduireMemoire('Synthèse des pièces;Fonction appelable uniquement depuis une pièce d''indice 0')+'W;O;O;O;','','');
  Result := TobAplat;
end;

function Genere_Tob_Client(sClient: string; TypeSortie: string = 'E'; sWhere: string = ''): Tob;
var
  Toblignes, TobAplat: Tob;
begin
  TobAplat := nil;
  MakeListes;
  TobLignes := Charge_Tob_Client(sClient, TypeSortie, sWhere);
  if Assigned(TobLignes) then
  begin
    Charge_Arborescence_Pieces(TobLignes, TypeSortie);
    TobAplat := Mets_A_Plat(TobLignes, False);
  end;
  Result := TobAplat;
end;

function Genere_Tob_Fournisseur(sFournisseur: string; TypeSortie: string = 'E'; sWhere: string = ''): Tob;
var
  Toblignes, TobAplat: Tob;
begin
  TobAplat := nil;
  MakeListes;
  TobLignes := Charge_Tob_Fournisseur(sFournisseur, TypeSortie, sWhere);
  if Assigned(toblignes) then
  begin
    Charge_Arborescence_Pieces(TobLignes, TypeSortie);
    TobAplat := Mets_A_Plat(TobLignes, False);
  end;
  Result := TobAplat;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Pascal Gautherot
Créé le ...... : 15/01/2003
Modifié le ... :   /  /
Description .. : "Registration" des procédures et fonctions implémentées
Suite ........ : dans cette unitée et susceptibles d'être appellées depuis les
Suite ........ : scripts.
Mots clefs ... :
*****************************************************************}

procedure RegisterProc;
begin
  RegisterAglProc('AglPiecesClient', FALSE, 1, AglPiecesClient);
  RegisterAglProc('AglPiecesFournisseur', FALSE, 1, AglPiecesFournisseur);
end;

initialization
  RegisterProc();
end.
