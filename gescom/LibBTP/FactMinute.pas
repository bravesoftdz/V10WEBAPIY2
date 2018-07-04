unit FactMinute;

interface
uses UTOB,HEnt1,
{$IFDEF V530}
  EdtEtat,
{$ELSE}
  EdtREtat,
{$ENDIF}
FactOuvrage, FactUtil,
SysUtils;

procedure ImprimeMinute (TobArticles,TOBPiece,TOBOUvrage,TOBTiers,TOBAffaire : TOB);

implementation

uses factvariante,
     ParamSoc;

var TOTPATOT,TOTPVTOT,TOTHRSTOT,TOTPRTOT : double;
    ImprimeVariante : Boolean;

procedure AddchampBase (TOBMinute,TOBPiece,TOBTiers,TOBAffaire : TOB);
begin
TOBMinute.AddChampSupValeur ('BMN_NATUREPIECEG',TOBPIece.getValue('GP_NATUREPIECEG'),false);
TOBMinute.AddChampSupValeur ('BMN_NUMPIECE',TOBPIece.getValue('GP_NUMERO'),false);
TOBMinute.AddChampSupValeur ('BMN_DATE',TOBPIece.getValue('GP_DATEPIECE'),false);
TOBMinute.AddChampSupValeur ('BMN_LIBELLEAFF',TOBAffaire.getValue('AFF_LIBELLE'),false);
TOBMinute.AddChampSupValeur ('BMN_LIBTIERS',TOBTiers.getValue('T_LIBELLE'),false);
TOBMinute.AddChampSupValeur ('BMN_REFINTERNE',TOBPIece.getValue('GP_REFINTERNE'),false);
TOBMinute.AddChampSupValeur ('BMN_REFEXTERNE',TOBPIece.getValue('GP_REFEXTERNE'),false);
end;

procedure AddchampLigne (TOBMinute: TOB);
begin
TOBMinute.AddChampSupValeur ('BMN_TYPELIGNE','',false);
TOBMinute.AddChampSupValeur ('BMN_TYPEARTICLE','',false);
TOBMinute.AddChampSupValeur ('BMN_CODEARTICLE','',false);
TOBMinute.AddChampSupValeur ('BMN_LIBELLE','',false);
TOBMinute.AddChampSupValeur ('BMN_QTEFACT',0,false);
TOBMinute.AddChampSupValeur ('BMN_UNITE','',false);
TOBMinute.AddChampSupValeur ('BMN_PUA',0,false);
TOBMinute.AddChampSupValeur ('BMN_MTDEBOURSE',0,false);
TOBMinute.AddChampSupValeur ('BMN_MTREVIENT',0,false);
TOBMinute.AddChampSupValeur ('BMN_NBHRS',0,false);
TOBMinute.AddChampSupValeur ('BMN_UNITEHRS','',false);
TOBMinute.AddChampSupValeur ('BMN_PUV',0,false);
TOBMinute.AddChampSupValeur ('BMN_MONTANT',0,false);
TOBMinute.AddChampSupValeur ('BMN_QTEDUDETAIL',1,false);
TOBMinute.AddChampSupValeur ('BMN_VARIANTE','-',false);
end;

procedure AjouteTotDocument (TOBPiece,TOBTiers,TOBAffaire,TOBMinute : TOB);
var TOBM : TOB;
begin
TOBM := TOB.Create ('MINUTE',TOBMinute,-1);
AddchampBase (TOBM,TOBPiece,TOBTiers,TOBAffaire);
AddchampLigne(TOBM);
TOBM.PutValue('BMN_TYPELIGNE',  'FMN');
TOBM.PutValue('BMN_TYPEARTICLE','COM');
TOBM.PutValue('BMN_MTDEBOURSE', TOTPATOT);
TOBM.PutValue('BMN_MTREVIENT',  TOTPRTOT);
TOBM.PutValue('BMN_MONTANT',    TOTPVTOT);
TOBM.PutValue('BMN_NBHRS',      TOTHRSTOT);
end;

procedure AjouteParagrapheMinute (TOBPiece,TOBTiers,TOBAffaire,TOBL,TOBMinute : TOB);
var TOBM : TOB;
begin
//
TOBM := TOB.Create ('MINUTE',TOBMinute,-1);
AddchampBase (TOBM,TOBPiece,TOBTiers,TOBAffaire);
AddchampLigne (TOBM);
TOBM.PutValue('BMN_TYPELIGNE','DP');
TOBM.PutValue('BMN_TYPEARTICLE','COM');
TOBM.PutValue('BMN_LIBELLE',TOBL.GetValue('GL_LIBELLE'));
end;

procedure AjouteFinParagrapheMinute (TOBPiece,TOBTiers,TOBAffaire,TOBL,TOBMinute : TOB);
var TOBM : TOB;
begin
TOBM := TOB.Create ('MINUTE',TOBMinute,-1);
AddchampBase (TOBM,TOBPiece,TOBTiers,TOBAffaire);
AddchampLigne (TOBM);
TOBM.PutValue('BMN_TYPELIGNE','TP');
TOBM.PutValue('BMN_LIBELLE',TOBL.GetValue('GL_LIBELLE'));
TOBM.PutValue('BMN_MTDEBOURSE',TOBL.GetValue('GL_MONTANTPA'));
TOBM.PutValue('BMN_MTREVIENT',TOBL.GetValue('GL_MONTANTPR'));
TOBM.PutValue('BMN_NBHRS',TOBL.GetValue('GL_TPSUNITAIRE'));
TOBM.PutValue('BMN_MONTANT',TOBL.GetValue('GL_MONTANTHTDEV'));
end;

procedure AjouteArticleMinute (TOBPiece,TOBL,TOBTiers,TOBAffaire,TOBArticles,TOBOuvrage,TOBMinute : TOB);
var TOBO,TOBM : TOB;
		Indice,niveau : integer;
begin

TOBM := TOB.Create ('MINUTE',TOBMinute,-1);

AddchampBase (TOBM,TOBPiece,TOBTiers,TOBAffaire);
AddchampLigne(TOBM);

TOBM.PutValue ('BMN_TYPELIGNE',   'ART');
TOBM.PutValue ('BMN_TYPEARTICLE', 'MAR');
TOBM.PutValue ('BMN_CODEARTICLE', TOBL.GetValue('GL_CODEARTICLE'));
TOBM.PutValue ('BMN_LIBELLE',     TOBL.GetValue('GL_LIBELLE'));
TOBM.PutValue ('BMN_UNITE',       TOBL.GetValue('GL_QUALIFQTEVTE'));
TOBM.PutValue ('BMN_QTEFACT',     TOBL.GetValue('GL_QTEFACT'));
TOBM.PutValue ('BMN_PUA',         TOBL.GetValue('GL_DPA'));
TOBM.PutValue ('BMN_MTDEBOURSE',  TOBL.GetValue('GL_MONTANTPA'));
TOBM.PutValue ('BMN_MTREVIENT',   TOBL.GetValue('GL_MONTANTPR'));
TOBM.PutValue ('BMN_NBHRS',Arrondi(TOBL.GetValue('GL_TPSUNITAIRE')*TOBL.GetValue('GL_QTEFACT'),V_PGI.OKDecQ));
TOBM.PutValue ('BMN_PUV',         TOBL.GetValue('GL_PUHTDEV'));
TOBM.PutValue ('BMN_MONTANT',     TOBL.GetValue('GL_MONTANTHTDEV'));
//
  //FV1 : 19/02/2018 - FS#2927 - TREUIL - Edition minute : En totalisation générale ne pas prendre en compte les variantes
  if IsVariante(TOBL) then
  Begin
    TOBMinute.PutValue('BMN_VARIANTE','X');
  end
  else
  begin
    TOTPATOT  := TOTPATOT  + TOBL.GetValue('GL_MONTANTPA');
    TOTPRTOT  := TOTPRTOT  + TOBL.GetValue('GL_MONTANTPR');
    TOTPVTOT  := TOTPVTOT  + TOBL.GetValue('GL_MONTANTHTDEV') ;
    TOTHRSTOT := TOTHRSTOT + Arrondi(TOBL.GetValue('GL_TPSUNITAIRE')*TOBL.GetValue('GL_QTEFACT'),V_PGI.OKDecQ);
  end;

end;

procedure AjouteOuvrageMinute (TOBPiece,TOBL,TOBTiers,TOBAffaire,TOBOuvrage,TOBMinute : TOB);
var Indice,niveau : integer;
    TOBO,TOBGO,TOBM : TOB;
    Qte : double;
begin

  if TOBL.GetValue('GL_INDICENOMEN') = 0 then Exit;
  TOBGO := TOBOuvrage.detail [TOBL.GetValue('GL_INDICENOMEN')-1];
  if TOBGO = nil then exit;
  TOBM := TOB.Create ('MINUTE',TOBMinute,-1);
  AddchampBase (TOBM,TOBPiece,TOBTiers,TOBAffaire);
  AddchampLigne (TOBM);

  TOBM.PutValue ('BMN_TYPELIGNE','OUV');
  TOBM.PutValue ('BMN_CODEARTICLE',TOBL.GetValue('GL_CODEARTICLE'));
  TOBM.PutValue ('BMN_LIBELLE',TOBL.GetValue('GL_LIBELLE'));
  TOBM.PutValue ('BMN_UNITE',TOBL.GetValue('GL_QUALIFQTEVTE'));
  TOBM.PutValue ('BMN_QTEFACT',TOBL.GetValue('GL_QTEFACT'));
  TOBM.PutValue ('BMN_PUA',TOBL.GetValue('GL_DPA'));
  TOBM.PutValue ('BMN_MTDEBOURSE',TOBL.GetValue('GL_MONTANTPA'));
  TOBM.PutValue ('BMN_MTREVIENT',TOBL.GetValue('GL_MONTANTPR'));
  TOBM.PutValue ('BMN_NBHRS',Arrondi(TOBL.GetValue('GL_TPSUNITAIRE')*TOBL.GetValue('GL_QTEFACT'),V_PGI.OKDecQ));
  TOBM.PutValue ('BMN_PUV',TOBL.GetValue('GL_PUHTDEV'));
  TOBM.PutValue ('BMN_MONTANT',TOBL.GetValue('GL_MONTANTHTDEV'));

  //FV1 : 19/02/2018 - FS#2927 - TREUIL - Edition minute : En totalisation générale ne pas prendre en compte les variantes
  If isvariante(TOBL) then
  Begin
    TOBM.PutValue('BMN_VARIANTE','X');
  end;

  if TOBGO.detail.count > 0 then
  begin
    TOBM.PutValue('BMN_QTEDUDETAIL',TOBGO.detail[0].getValue('BLO_QTEDUDETAIL'));
    for Indice := 0 to TOBGO.detail.Count -1 do
    begin
      TOBO := TOBGO.detail[Indice];
      TOBM := TOB.Create ('MINUTE',TOBMinute,-1);
      AddchampBase (TOBM,TOBPiece,TOBTiers,TOBAffaire);
      AddchampLigne (TOBM);
      TOBM.PutValue ('BMN_TYPELIGNE','ART');
      TOBM.PutValue ('BMN_TYPEARTICLE','MAR');
      TOBM.PutValue ('BMN_CODEARTICLE',TOBO.GetValue ('BLO_CODEARTICLE'));
      TOBM.PutValue ('BMN_LIBELLE',TOBO.GetValue ('BLO_LIBELLE'));
      TOBM.PutValue ('BMN_QTEFACT',TOBO.GetValue ('BLO_QTEFACT'));
      TOBM.PutValue ('BMN_UNITE',TOBO.GetValue('BLO_QUALIFQTEVTE'));
      TOBM.PutValue ('BMN_PUA',TOBO.GetValue('BLO_DPA'));
      TOBM.PutValue ('BMN_MTDEBOURSE',TOBO.GetValue('BLO_MONTANTPA'));
      TOBM.PutValue ('BMN_MTREVIENT',TOBO.GetValue('BLO_MONTANTPR'));
      TOBM.PutValue ('BMN_NBHRS',Arrondi(TOBO.GetValue('BLO_TPSUNITAIRE')*TOBO.GetValue('BLO_QTEFACT'),V_PGI.OKDecQ));
      TOBM.PutValue ('BMN_PUV',TOBO.GetValue('BLO_PUHTDEV'));
      TOBM.PutValue ('BMN_MONTANT',TOBO.GetValue('BLO_MONTANTHTDEV'));
      // VARIANTE
      //FV1 : 19/02/2018 - FS#2927 - TREUIL - Edition minute : En totalisation générale ne pas prendre en compte les variantes
      if IsVariante(TOBL) then
      Begin
        TOBM.PutValue('BMN_VARIANTE','X');
      end;
    end;
  end;

  //FV1 : 19/02/2018 - FS#2927 - TREUIL - Edition minute : En totalisation générale ne pas prendre en compte les variantes
  if not IsVariante(TOBL) then
  begin
    TOTPATOT := TOTPATOT + TOBL.GetValue('GL_MONTANTPA');
    TOTPRTOT := TOTPRTOT + TOBL.GetValue('GL_MONTANTPR');
    TOTPVTOT := TOTPVTOT + TOBL.GetValue('GL_MONTANTHTDEV') ;
    TOTHRSTOT := TOTHRSTOT + Arrondi(TOBL.GetValue('GL_TPSUNITAIRE')*TOBL.GetValue('GL_QTEFACT'),V_PGI.OKDecQ);
  end;

  TOBM := TOB.Create ('MINUTE',TOBMinute,-1);
  AddchampBase (TOBM,TOBPiece,TOBTiers,TOBAffaire);
  AddchampLigne (TOBM);

  TOBM.PutValue ('BMN_TYPELIGNE','FOV');

end;

procedure ImprimeMinute (TobArticles,TOBPiece,TOBOUvrage,TOBTiers,TOBAffaire : TOB);
var Indice      : integer;
    TOBMinute   : TOB;
    TOBL        : TOB;
    TypeArticle : string;
    Modele      : string;
begin

  ImprimeVariante := GetParamSocSecur('SO_IMPVARIANTEMINUTE', False);

  TOBMinute := TOB.Create ('DESCRIPTIFMINUTES',nil,-1);

  TRY
    TOTPATOT := 0;
    TOTPRTOT := 0;
    TOTPVTOT := 0;
    TOTHRSTOT := 0;
    for indice := 0 to TOBPiece.detail.count -1 do
    begin
      TOBL := TOBPiece.detail[Indice];
      TypeArticle := TOBL.GetValue('GL_TYPEARTICLE');
      if copy(TOBL.GetValue('GL_TYPELIGNE'),1,2)='TP' then
      begin
        AjouteFinParagrapheMinute (TOBPiece,TOBTiers,TOBAffaire,TOBL,TOBMinute);
      end;
      if copy(TOBL.GetValue('GL_TYPELIGNE'),1,2)='DP' then
      begin
        AjouteParagrapheMinute (TOBPiece,TOBTiers,TOBAffaire,TOBL,TOBMinute);
      end;
      // VARIANTE
      if ImprimeVariante then
      begin
        if copy(TOBL.GetValue('GL_TYPELIGNE'),1,2)='TV' then
        begin
          AjouteFinParagrapheMinute (TOBPiece,TOBTiers,TOBAffaire,TOBL,TOBMinute);
        end;
        if copy(TOBL.GetValue('GL_TYPELIGNE'),1,2)='DV' then
        begin
          AjouteParagrapheMinute (TOBPiece,TOBTiers,TOBAffaire,TOBL,TOBMinute);
        end;      
        if ((TypeArticle='MAR') or (TypeArticle = 'PRE') or (TypeArticle = 'POU') or (TypeArticle = 'ARP')) and (not IsSousDetail(TOBL)) and (not IsCommentaire(TOBL)) then
        begin
          AjouteArticleMinute (TOBPiece,TOBL,TOBTiers,TOBAffaire,TOBArticles,TOBOuvrage,TOBMinute);
        end;
        if (TypeArticle='OUV') then
        begin
          AjouteOuvrageMinute (TOBPiece,TOBL,TOBTiers,TOBAffaire,TOBOuvrage,TOBMinute);
        end;
      end
      else
      begin
        if ((TypeArticle='MAR') or (TypeArticle = 'PRE') or (TypeArticle = 'POU') or (TypeArticle = 'ARP')) and (not IsVariante(TOBL)) and (not IsSousDetail(TOBL)) and (not IsCommentaire(TOBL)) then
        begin
          AjouteArticleMinute (TOBPiece,TOBL,TOBTiers,TOBAffaire,TOBArticles,TOBOuvrage,TOBMinute);
        end;
        if (TypeArticle='OUV') and (not IsVariante(TOBL)) then
        begin
          AjouteOuvrageMinute (TOBPiece,TOBL,TOBTiers,TOBAffaire,TOBOuvrage,TOBMinute);
        end;
      end;
    end;

    AjouteTotDocument(TOBPiece,TOBTiers,TOBAffaire,TOBMinute);

    Modele := GetParamSocSecur('SO_ETATMINUTE','BMN');

    lanceEtatTob ('E','GPJ',Modele,TOBMinute,true,false,false,nil,'','Edition de la minute',false);
  FINALLY
    TOBMinute.free;
  END;

end;

end.
