{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : GCVENTES ()
Mots clefs ... : TOF;GCVENTES
*****************************************************************}
Unit AchatsCubeTof ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     sysutils,
     ComCtrls,
     Forms,
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,
     UTOF,
     HTB97,
     UtilGC,

     vierge,
{$IFNDEF EAGLCLIENT}
     db,
     FE_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     uTob,
{$ENDIF}
     Windows;

Type
  TOF_GCACHATS = Class (TOF)

    procedure OnArgument (S : String ) ; override ;

  Private
    NATUREPIECEG : THMultiValComboBox;
    FAMILLENIV1  : THMultiValComboBox;
    FAMILLENIV2  : THMultiValComboBox;
    FAMILLENIV3  : THMultiValComboBox;
    //
    LFAMILLENIV1 : THLabel;
    LFAMILLENIV2 : THLabel;
    LFAMILLENIV3 : THLabel;
    LLIBRETIERS1 : THLabel;
    LLIBRETIERS2 : THLabel;
    LLIBRETIERS3 : THLabel;
    //
    LIBRETIERS1   : THEdit;
    LIBRETIERS2   : THEdit;
    LIBRETIERS3   : THEdit;
    //
    TIERS         : THEdit;
    ARTICLE       : THEdit;
    //
    ETABLISSEMENT : THValComboBox;
    DOMAINE       : THValComboBox;
    //
    procedure ArticleOnElipsisClick(Sender: Tobject);
    procedure GetObject;
    procedure TiersOnElipsisClick(Sender: Tobject);

  end;

Implementation

uses affaireutil,
     Paramsoc,
     Dialogs,
     FactUtil,
     GerePiece,
     FileCtrl,
     FactComm;



procedure TOF_GCACHATS.OnArgument (S : String ) ;
Var i         : integer;
    AuMoinsUn : boolean;
begin
  Inherited ;

  GetObject;

  AuMoinsUn:=False;

  For i:=1 to 10 do if i<10 then
    AuMoinsUn:=(ChangeLibre2('LLIBRETIERS'+intToStr(i),TForm(Ecran)) or AuMoinsUn)
  else
    AuMoinsUn:=(ChangeLibre2('LLIBRETIERSA',TForm(Ecran))or AuMoinsUn);

  MajChampsLibresArticle(TForm(Ecran),'GL',(not AuMoinsUn),'PCOMPLEMENTS','PCOMPLEMENTS');

  if Etablissement <> Nil then PositionneEtabUser(Etablissement) ;

  if Domaine <> Nil then PositionneDomaineUser(Domaine) ;

  NATUREPIECEG.text := 'BLF;';

  LFAMILLENIV1.Caption := RechDom('GCLIBFAMILLE','LF1',False);
  LFAMILLENIV2.Caption := RechDom('GCLIBFAMILLE','LF2',False);
  LFAMILLENIV3.Caption := RechDom('GCLIBFAMILLE','LF3',False);

  LLIBRETIERS1.Caption := RechDom('GCZONELIBRE','FT1',False);
  if copy(LLIBRETIERS1.Caption,1,2) = '.-' then
  begin
    LLIBRETIERS1.visible := False;
    LIBRETIERS1.visible  := False;
  end;

  LLIBRETIERS2.Caption := RechDom('GCZONELIBRE','FT2',False);
  if copy(LLIBRETIERS2.Caption,1,2) ='.-' then
  begin
    LLIBRETIERS2.visible:=False;
    LIBRETIERS2.visible:=False;
  end;

  LLIBRETIERS3.Caption:=RechDom('GCZONELIBRE','FT3',False);
  if copy(LLIBRETIERS3.Caption,1,2) ='.-' then
  begin
    LLIBRETIERS3.visible:=False;
    LIBRETIERS3.visible:=False;
  end;

end ;

procedure TOF_GCACHATS.GetObject;
begin

  NaturePieceg  := THMultiValComboBox(GetControl('GL_NATUREPIECEG'));

  LIBRETIERS1   := THEdit(GetControl('GP_LIBRETIERS1'));
  LIBRETIERS2   := THEdit(GetControl('GP_LIBRETIERS2'));
  LIBRETIERS3   := THEdit(GetControl('GP_LIBRETIERS3'));
  //
  LLIBRETIERS2  := THLabel(GetControl('TGP_LIBRETIERS1'));
  LLIBRETIERS2  := THLabel(GetControl('TGP_LIBRETIERS1'));
  LLIBRETIERS2  := THLabel(GetControl('TGP_LIBRETIERS1'));
  //
  LFAMILLENIV1  := THLabel(GetControl('TGP_FAMILLENIV1'));
  LFAMILLENIV2  := THLabel(GetControl('TGP_FAMILLENIV2'));
  LFAMILLENIV3  := THLabel(GetControl('TGP_FAMILLENIV3'));
  //
  FAMILLENIV1   := THMultiValComboBox(GetControl('GL_FAMILLENIV1'));
  FAMILLENIV2   := THMultiValComboBox(GetControl('GL_FAMILLENIV2'));
  FAMILLENIV3   := THMultiValComboBox(GetControl('GL_FAMILLENIV3'));
  //
  ETABLISSEMENT := THValComboBox(GetControl('GL_ETABLISSEMENT'));
  DOMAINE       := THValComboBox(GetControl('GL_DOMAINE'));
  //
  TIERS         := THEdit(GetControl('GL_TIERS'));
  Tiers.OnElipsisClick := TiersOnElipsisClick;
  //
  Article       := THEdit(GetControl('GL_ARTICLE'));
  Article.OnElipsisClick := ArticleOnElipsisClick;
end;

Procedure TOF_GCACHATS.TiersOnElipsisClick(Sender : Tobject);
Var StChamps  : string;
begin

  StChamps  := Tiers.Text;

  Tiers.Text := AGLLanceFiche('GC','GCFOURNISSEUR_MUL','T_TIERS=' + Tiers.text +';T_NATUREAUXI=FOU','','SELECTION');

end;

Procedure TOF_GCACHATS.ArticleOnElipsisClick(Sender : Tobject);
Var StChamps  : string;
    SWhere    : string;
begin

	sWhere := ' AND ((GA_TYPEARTICLE="MAR") AND GA_TENUESTOCK="X")';

  StChamps := Article.Text;

  if Article.Text <> '' then
    sWhere := 'GA_CODEARTICLE=' + Trim(Copy(Article.text, 1, 18)) + ';XX_WHERE=' + sWhere
  else
    sWhere := 'XX_WHERE=' + sWhere;

	Article.text := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', sWhere+';RECHERCHEARTICLE');
  Article.text := Trim(Copy(Article.text, 1, 18));

end;


Initialization
  registerclasses ( [ TOF_GCACHATS ] ) ;
end.

