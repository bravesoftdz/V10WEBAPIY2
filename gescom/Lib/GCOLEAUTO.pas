unit GCOLEAUTO;

interface

uses ComObj, ActiveX, HEnt1, HCtrls,
  {$IFDEF EAGLCLIENT}
  UTob, eFichList,
  {$ELSE}
  DB, DBtables, FichList,
  {$ENDIF}
  StdVcl, OLEAuto,EntGC;
type
  TGCWINDOW = class(TAutoObject)
  automated
    function GCCumulPiece(const NomChamp, Natures, Etabs, Devises, Repres, Depots: string; DateDeb, DateFin: TDateTime; const Arg: string): Variant;
    function GCCumulLigne(const NomChamp, Natures, Etabs, Devises, Repres, Depots: string; DateDeb, DateFin: TDateTime; const Arg: string): Variant;
    function GCRechDOMOLE(const TT, Code: string): Variant;
    function GCCumulArticle(const NomChamp, Natures, Repres, FM1, FM2, FM3, Types, Statuts: string; DateDeb, DateFin: TDateTime; const Arg: string): Variant;
    procedure GCZoom(const Table: string; const Qui: string);
    function GCGetIDCourant: Variant;
    function GCTransformePiece(OldNat: Variant; NewNat: Variant; Numero: Variant;
      Souche: Variant; Indice: Variant; NewDate: TDateTime): Variant;
    function GCGetSQL(const StSQL: string): Variant;
    function GCCumulStat(const NomChamp, Natures, Repres, FM1, FM2, FM3: string;
      librearticle1, librearticle2, librearticle3, librearticle4, librearticle5,
      librearticle6, librearticle7, librearticle8, librearticle9, librearticlea: string;
      libretiers1, libretiers2, libretiers3, libretiers4, libretiers5,
      libretiers6, libretiers7, libretiers8, libretiers9, libretiersa: string;
      //           librepiece1,librepiece2,librepiece3,libreaffaire1,libreaffaire2,libreaffaire3 : string;
      DateDeb, DateFin: TDateTime; const Arg: string): Variant;
    function MODEMulArticle: Variant;
    function LibelleDimension(InfoDim: string): Variant;
    function MODEMultiSQL(StSQL: String; Multi:Variant): Variant;
  protected
    procedure OnTheLast(var shutdown: boolean);
  public
    constructor Create; override;
  end;

procedure RegisterGCPGI;

implementation

uses {ComServ,}  Dialogs, Forms, UTom, UtomLiensOle, FactGrp, UtilArticle;

procedure RegisterGCPGI; // appelé dans la phase initialization
const
  AutoClassInfo: TAutoClassInfo = (
    AutoClass: TGCWINDOW;
    ProgID: 'CGS5.GCWINDOW';
    // ATTENTION : CE NUMERO DE CLASSE ID EST A CHANGER
    // IL DOIT LIE A UNE APPLICATION ET A UNE SEULE
    // POUR CALCULER UN GUID UNIQUE, FAIRE DANS DELPHI :
    // CTRL+SHIFT+G
    {$IFDEF CCS3}
    {$IFDEF EAGLCLIENT}
    ClassID: '{877C3B0F-A61E-4A62-9571-322A1D87A73F}';
    {$ELSE}
    ClassID: '{812FDA01-EF30-450E-9E55-AB8C030ABBBF}';
    {$ENDIF}
    {$ELSE}
    {$IFDEF EAGLCLIENT}
    ClassID: '{604E14C8-5D70-4F06-95FF-117F01934D64}';
    {$ELSE}
    ClassID: '{7FED4FE5-4B31-464B-BD34-F5D0D62D7FB6}';
    {$ENDIF}
    {$ENDIF}
    Description: 'Cegid GC PGI';
    Instancing: acMultiInstance);
begin
  Automation.RegisterClass(AutoClassInfo);
end;

procedure TGCWINDOW.OnTheLast(var shutdown: boolean);
begin
  ShutDown := FALSE; // pour ne pas fermer l'application automatiquement lorsqu'il n'y a plus de réference à l'objet.
end;

constructor TGCWINDOW.create;
begin
  inherited;
  Automation.OnLastRelease := OnTheLast;
end;

function DecodeToken(StOrig, Champ: string; EstDans: Boolean = False): string;
var ii: integer;
  St, StLoc, StRes, estPas: string;
begin
  Result := '';
  St := StOrig;
  StRes := '';
  ii := 0;
  estPas := '';
  if EstDans then
  begin
    if st[1] = '#' then
    begin
      estPas := ' NOT ';
      Delete(St, 1, 1);
    end;
    StRes := champ + estPas + ' IN (';
    repeat
      StLoc := ReadTokenSt(St);
      if StLoc <> '' then
      begin
        StRes := StRes + '"' + StLoc + '",';
        Inc(ii);
      end;
    until ((St = '') or (StLoc = ''));
    if ii > 0 then
    begin
      Delete(StRes, Length(StRes), 1); //supprime la virgule
      Result := Stres + ') '
    end else Result := '';
  end else
  begin
    repeat
      StLoc := ReadTokenSt(St);
      if StLoc <> '' then
      begin
        StRes := StRes + Champ + '="' + StLoc + '" OR ';
        Inc(ii);
      end;
    until ((St = '') or (StLoc = ''));
    if StRes <> '' then Delete(StRes, Length(StRes) - 3, 4);
    if ii > 1 then StRes := '(' + Stres + ')';
    Result := StRes;
  end;
end;

function CompleteSQLOLE(StVals, Champ: string; var PasWhere: Boolean; EstDans: Boolean = False): string;
var StLoc: string;
begin
  Result := '';
  if StVals <> '' then
  begin
    StLoc := DecodeToken(StVals, Champ, EstDans);
    if PasWhere then
    begin
      Result := ' WHERE ';
      PasWhere := False;
    end else Result := ' AND ';
    Result := Result + StLoc;
  end;
end;

procedure VarSQLWhere(var SQL: string; var PasWhere: boolean);
begin
  if PasWhere then
  begin
    SQL := SQL + ' WHERE ';
    PasWhere := False;
  end else SQL := SQL + ' AND ';
end;

function TGCWINDOW.GCRechDOMOLE(const TT, Code: string): Variant;
begin
  Result := '#Erreur : Non connecté';
  if not V_PGI.OKOuvert then Exit;
  Result := RechDom(TT, Code, False);
end;

function TGCWINDOW.GCCumulPiece(const NomChamp, Natures, Etabs, Devises, Repres, Depots: string; DateDeb, DateFin: TDateTime; const Arg: string): Variant;
var Q: TQuery;
  SQL: string;
  PasWhere: Boolean;
begin
  Result := '#Erreur : Non connecté';
  if not V_PGI.OKOuvert then Exit;
  if VH_GC.GCIfDefCEGID then
    if (datedeb < 2) or (dateFin < 2) then
    begin
      Result := '#Erreur : Dates erronées';
      Exit
    end;
  SQL := 'Select SUM(' + NomChamp + ') FROM PIECE';
  PasWhere := True;
  SQL := SQL + CompleteSQLOLE(Natures, 'GP_NATUREPIECEG', PasWhere);
  SQL := SQL + CompleteSQLOLE(Etabs, 'GP_ETABLISSEMENT', PasWhere);
  SQL := SQL + CompleteSQLOLE(Devises, 'GP_DEVISE', PasWhere);
  SQL := SQL + CompleteSQLOLE(Repres, 'GP_REPRESENTANT', PasWhere);
  SQL := SQL + CompleteSQLOLE(Depots, 'GP_DEPOT', PasWhere);
  if Arg <> '' then
  begin
    VarSQLWhere(SQL, PasWhere);
    SQL := SQL + Arg;
  end;
  if DateDeb > 2 then
  begin
    VarSQLWhere(SQL, PasWhere);
    SQL := SQL + 'GP_DATEPIECE>="' + USDateTime(DateDeb) + '"';
  end;
  if DateFin > 2 then
  begin
    VarSQLWhere(SQL, PasWhere);
    SQL := SQL + 'GP_DATEPIECE<="' + USDateTime(DateFin) + '"';
  end;
  Q := OpenSQL(SQL, True);
  if not Q.EOF then Result := Q.Fields[0].AsVariant else Result := 0;
  Ferme(Q);
end;

function TGCWINDOW.GCCumulLigne(const NomChamp, Natures, Etabs, Devises, Repres, Depots: string; DateDeb, DateFin: TDateTime; const Arg: string): Variant;
var Q: TQuery;
  SQL: string;
  PasWhere: Boolean;
begin
  Result := '#Erreur : Non connecté';
  if not V_PGI.OKOuvert then Exit;
  if VH_GC.GCIfDefCEGID then
    if (datedeb < 2) or (dateFin < 2) then
    begin
      Result := '#Erreur : Dates erronées';
      Exit
    end;
  SQL := 'Select SUM(' + NomChamp + ') FROM LIGNE';
  PasWhere := True;
  SQL := SQL + CompleteSQLOLE(Natures, 'GL_NATUREPIECEG', PasWhere);
  SQL := SQL + CompleteSQLOLE(Etabs, 'GL_ETABLISSEMENT', PasWhere);
  SQL := SQL + CompleteSQLOLE(Devises, 'GL_DEVISE', PasWhere);
  SQL := SQL + CompleteSQLOLE(Repres, 'GL_REPRESENTANT', PasWhere);
  SQL := SQL + CompleteSQLOLE(Depots, 'GL_DEPOT', PasWhere);
  if Arg <> '' then
  begin
    VarSQLWhere(SQL, PasWhere);
    SQL := SQL + Arg;
  end;
  if DateDeb > 2 then
  begin
    VarSQLWhere(SQL, PasWhere);
    SQL := SQL + 'GL_DATEPIECE>="' + USDateTime(DateDeb) + '"';
  end;
  if DateFin > 2 then
  begin
    VarSQLWhere(SQL, PasWhere);
    SQL := SQL + 'GL_DATEPIECE<="' + USDateTime(DateFin) + '"';
  end;
  if PasWhere then SQL := SQL + ' WHERE GL_TYPELIGNE="ART"' else SQL := SQL + ' AND GL_TYPELIGNE="ART"';
  Q := OpenSQL(SQL, True);
  if not Q.EOF then Result := Q.Fields[0].AsVariant else Result := 0;
  Ferme(Q);
end;

function TGCWINDOW.GCCumulArticle(const NomChamp, Natures, Repres, FM1, FM2, FM3, Types, Statuts: string; DateDeb, DateFin: TDateTime; const Arg: string):
  Variant;
var Q: TQuery;
  SQL: string;
  PasWhere: Boolean;
begin
  Result := '#Erreur : Non connecté';
  if not V_PGI.OKOuvert then Exit;
  if VH_GC.GCIfDefCEGID then
    if (datedeb < 2) or (dateFin < 2) then
    begin
      Result := '#Erreur : Dates erronées';
      Exit
    end;
  SQL := 'Select SUM(' + NomChamp + ') FROM LIGNE LEFT JOIN ARTICLE ON GL_ARTICLE=GA_ARTICLE';
  PasWhere := True;
  SQL := SQL + CompleteSQLOLE(Natures, 'GL_NATUREPIECEG', PasWhere);
  SQL := SQL + CompleteSQLOLE(Repres, 'GL_REPRESENTANT', PasWhere);
  SQL := SQL + CompleteSQLOLE(FM1, 'GA_FAMILLENIV1', PasWhere);
  SQL := SQL + CompleteSQLOLE(FM2, 'GA_FAMILLENIV2', PasWhere);
  SQL := SQL + CompleteSQLOLE(FM3, 'GA_FAMILLENIV3', PasWhere);
  SQL := SQL + CompleteSQLOLE(Types, 'GA_TYPEARTICLE', PasWhere);
  SQL := SQL + CompleteSQLOLE(Statuts, 'GA_STATUTART', PasWhere);
  if Arg <> '' then
  begin
    VarSQLWhere(SQL, PasWhere);
    SQL := SQL + Arg;
  end;
  if DateDeb > 2 then
  begin
    VarSQLWhere(SQL, PasWhere);
    SQL := SQL + 'GL_DATEPIECE>="' + USDateTime(DateDeb) + '"';
  end;
  if DateFin > 2 then
  begin
    VarSQLWhere(SQL, PasWhere);
    SQL := SQL + 'GL_DATEPIECE<="' + USDateTime(DateFin) + '"';
  end;
  if PasWhere then SQL := SQL + ' WHERE GL_TYPELIGNE="ART"' else SQL := SQL + ' AND GL_TYPELIGNE="ART"';
  Q := OpenSQL(SQL, True);
  if not Q.EOF then Result := Q.Fields[0].AsVariant else Result := 0;
  Ferme(Q);
end;

function TGCWINDOW.GCGetSQL(const StSQL: string): Variant;
var Q: tQuery;
begin
  Result := '#Erreur : Non connecté';
  if not V_PGI.OKOuvert then Exit;
  Q := OpenSQL(StSQL, TRUE);
  if not Q.Eof then Result := Q.Fields[0].AsVariant else Result := '#Erreur : la requète ne renvoie aucun résultat';
  Ferme(Q);
end;

function TGCWINDOW.GCCumulStat(const NomChamp, Natures, Repres, FM1, FM2, FM3: string;
  librearticle1, librearticle2, librearticle3, librearticle4, librearticle5,
  librearticle6, librearticle7, librearticle8, librearticle9, librearticlea: string;
  libretiers1, libretiers2, libretiers3, libretiers4, libretiers5,
  libretiers6, libretiers7, libretiers8, libretiers9, libretiersa: string;
  //           librepiece1,librepiece2,librepiece3,libreaffaire1,libreaffaire2,libreaffaire3 : string;
  DateDeb, DateFin: TDateTime; const Arg: string): Variant;
var Q: TQuery;
  SQL: string;
  PasWhere: Boolean;
begin
  Result := '#Erreur : Non connecté';
  if not V_PGI.OKOuvert then Exit;
  if VH_GC.GCIfDefCEGID then
    if (datedeb < 2) or (dateFin < 2) then
    begin
      Result := '#Erreur : Dates erronées';
      Exit
    end;
  SQL := 'Select SUM(' + NomChamp + ') FROM LIGNE LEFT JOIN PIECE ON LIEN(LIGNE,PIECE) ';
  PasWhere := True;
  SQL := SQL + CompleteSQLOLE(Natures, 'GL_NATUREPIECEG', PasWhere);
  SQL := SQL + CompleteSQLOLE(Repres, 'GL_REPRESENTANT', PasWhere);
  SQL := SQL + CompleteSQLOLE(FM1, 'GL_FAMILLENIV1', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(FM2, 'GL_FAMILLENIV2', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(FM3, 'GL_FAMILLENIV3', PasWhere, True);
  //SQL:=SQL+CompleteSQLOLE(types,'GL_TYPEARTICLE',PasWhere) ;
  SQL := SQL + CompleteSQLOLE(librearticle1, 'GL_LIBREART1', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(librearticle2, 'GL_LIBREART2', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(librearticle3, 'GL_LIBREART3', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(librearticle4, 'GL_LIBREART4', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(librearticle5, 'GL_LIBREART5', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(librearticle6, 'GL_LIBREART6', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(librearticle7, 'GL_LIBREART7', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(librearticle8, 'GL_LIBREART8', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(librearticle9, 'GL_LIBREART9', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(librearticleA, 'GL_LIBREARTA', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(libretiers1, 'GP_LIBRETIERS1', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(libretiers2, 'GP_LIBRETIERS2', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(libretiers3, 'GP_LIBRETIERS3', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(libretiers4, 'GP_LIBRETIERS4', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(libretiers5, 'GP_LIBRETIERS5', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(libretiers6, 'GP_LIBRETIERS6', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(libretiers7, 'GP_LIBRETIERS7', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(libretiers8, 'GP_LIBRETIERS8', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(libretiers9, 'GP_LIBRETIERS9', PasWhere, True);
  SQL := SQL + CompleteSQLOLE(libretiersA, 'GP_LIBRETIERSA', PasWhere, True);
  {SQL:=SQL+CompleteSQLOLE(librepiece1,'GL_LIBREPIECE1',PasWhere,True) ;
  SQL:=SQL+CompleteSQLOLE(librepiece2,'GL_LIBREPIECE2',PasWhere,True) ;
  SQL:=SQL+CompleteSQLOLE(librepiece3,'GL_LIBREPIECE3',PasWhere,True) ;
  SQL:=SQL+CompleteSQLOLE(libreaffaire1,'GL_LIBREAFF1',PasWhere,True) ;
  SQL:=SQL+CompleteSQLOLE(libreaffaire2,'GL_LIBREAFF2',PasWhere,True) ;
  SQL:=SQL+CompleteSQLOLE(libreaffaire3,'GL_LIBREAFF3',PasWhere,True) ; }
  if Arg <> '' then
  begin
    VarSQLWhere(SQL, PasWhere);
    SQL := SQL + Arg;
  end;
  if DateDeb > 2 then
  begin
    VarSQLWhere(SQL, PasWhere);
    SQL := SQL + 'GL_DATEPIECE>="' + USDateTime(DateDeb) + '"';
  end;
  if DateFin > 2 then
  begin
    VarSQLWhere(SQL, PasWhere);
    SQL := SQL + 'GL_DATEPIECE<="' + USDateTime(DateFin) + '"';
  end;
  if PasWhere then SQL := SQL + ' WHERE GL_TYPELIGNE="ART"' else SQL := SQL + ' AND GL_TYPELIGNE="ART"';
  Q := OpenSQL(SQL, True);
  if (not Q.EOF) and (vartype(Q.Fields[0].AsVariant) = varDouble)
    then Result := Q.Fields[0].AsVariant
  else result := 0;
  Ferme(Q);
end;

procedure TGCWINDOW.GCZoom(const Table: string; const Qui: string);
var SQL, Champ: string;
begin
  if Table = 'COMMERCIAL' then
  begin
    Champ := 'GCL_COMMERCIAL';
  end;
  SQL := 'SELECT ' + Champ + ' from ' + Table + ' WHERE ' + Champ + '="' + Qui + '"';
  if ExisteSQL(SQL) then
  begin
    InitZoomOLE;
    if Table = 'COMMERCIAL' then V_PGI.DispatchTT(9, taConsult, Qui, '', '');
    FinZoomOLE;
  end;
end;

function TGCWINDOW.GCGetIDCourant: Variant;
var F: TForm;
  OM: TOM;
begin
  result := '';
  F := TForm(Screen.ActiveForm);
  if (F is TFFicheListe) then OM := TFFicheListe(F).OM else exit;
  if (OM is TOM_LiensOle) then result := TOM_LiensOle(OM).GetIDCourant else exit;
end;

function TGCWINDOW.GCTransformePiece(OldNat: Variant; NewNat: Variant; Numero: Variant;
  Souche: Variant; Indice: Variant; NewDate: TDateTime): Variant;
begin
  Result := TransfoBatchPiece(OldNat, NewNat, Souche, Numero, Indice, NewDate);
end;

{initialization
  TAutoObjectFactory.Create(ComServer, TGCWINDOW, Class_GCWINDOW, ciMultiInstance, tmApartment);  }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier Tarcy
Créé le ...... : 18/09/2003
Modifié le ... : 18/09/2003
Description .. : Ouvre le mul de recherche article
Mots clefs ... : RECHERCHE;ARTICLE,OLE
*****************************************************************}
function TGCWINDOW.MODEMulArticle: Variant;
begin
  InitZoomOLE;
  Result := GetArticleRecherche_Disp(nil, '', 'FFO', '', True);
  FinZoomOLE;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier Tarcy
Créé le ...... : 18/09/2003
Modifié le ... : 18/09/2003
Description .. : Retourne les libellés de la dimension et du code article 
Suite ........ : passé en paramètres, sous forme tokenisée
Mots clefs ... : DIMENSION;OLE
*****************************************************************}
function TGCWINDOW.LibelleDimension(InfoDim: string): Variant;
var CodeArticle, iDim, stSQL: string;
begin
  Result := '';
  CodeArticle := ReadTokenSt(InfoDim);
  iDim := InfoDim;
  StSQL := 'select distinct gdi_libelle,gdi_rang from article ' +
    'left join dimension on gdi_typedim="DI' + iDim + '" and gdi_grilledim= ga_grilledim' + iDim + ' and gdi_codedim= ga_codedim' + iDim + ' ' +
    'where ga_statutart="DIM" and ga_codearticle="' + CodeArticle + '" ' +
    'order by gdi_rang';
  Result := MODEMultiSQL (StSQL,True) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier Tarcy
Créé le ...... : 19/09/2003
Modifié le ... : 19/09/2003
Description .. : renvoi le résultat d'une reqête SQL :
Suite ........ : -soit le 1er champ
Suite ........ : -soit tous les champs concaténés
Mots clefs ... : OLE;SQL
*****************************************************************}
function TGCWINDOW.MODEMultiSQL(StSQL: String; Multi:Variant): Variant;
var Q : TQuery;
begin
  Result := '';
  Q := OpenSQL(stSQL, True);
  if Multi then
  begin
    while not Q.EOF do
    begin
      if Result <> '' then
        Result := Result + ';' + Q.Fields[0].AsVariant else
        Result := Result + Q.Fields[0].AsVariant;
      Q.Next;
    end;
  end else
  Result := Q.Fields[0].AsVariant ;
  Ferme(Q);
end;

end.
