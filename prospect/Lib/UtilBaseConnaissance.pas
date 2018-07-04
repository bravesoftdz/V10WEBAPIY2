unit UtilBaseConnaissance;

interface
uses
{$IFDEF EAGLCLIENT}
      MaineAGL
{$ELSE}
      Fe_Main
{$ENDIF}
      ,Sysutils, HCtrls, HEnt1, HMsgBox
      ,M3FP
      ;

function LanceTheGoodMul(RATTACHEMENT1,RATTACHEMENT2,RATTACHEMENT3,RATTACHEMENT : THVALCOMBOBOX;
  Numchamp : integer) : string;

Const	Mul_BaseConnaissance : array[1..7] of string  = (
          {1} 'RBARTMUL_SELECT',
          {2} 'RBARTPARC_SELECT',
          {3} 'RBOPERATIO_SELECT',
          {4} 'RBRESSOURC_SELECT',
          {5} 'RBTIERS_SELECT',
          {6} 'RBCOMPET_SELECT',
          {7} 'RBTYPEACT_SELECT'
        );

    max_rattachement : integer = 7;
    Colonnes_BC : array[1..7] of string  = (
          {1} 'GA_LIBELLE;GA_ARTICLE;W;WARTICLE_FIC',
          {2} 'WAP_LIBELLE;WAP_ARTICLEPARC;W;WARTICLE_PARC',
          {3} 'ROP_LIBELLE;ROP_OPERATION;RT;RTOPERATIONS',
          {4} 'ARS_LIBELLE;ARS_RESSOURCE;AFF;RESSOURCE',
          {5} 'T_LIBELLE;T_TIERS;GC;GCTIERS',
          {6} 'PCO_LIBELLE;PCO_COMPETENCE;PAY;PGRHCOMPETENCES',
          {7} 'RPA_LIBELLE;RPA_TYPEACTION;RT;RTTYPEACTIONS'
        );

    NumFicheArt : integer = 1;
    NumFicheArtParc : integer = 2;
    NumFicheOpe : integer = 3;
    NumFicheRes : integer = 4;
    NumFichetiers : integer = 5;
    NumFicheCompetence : integer = 6;
    NumFicheTypeAction : integer = 7;
    PrefixeArt : string = 'GA';
    PrefixeArtParc : string = 'WAP';
    PrefixeOpe : string = 'ROP';
    PrefixeRes : string = 'ARS';
    Prefixetiers : string = 'T';
    PrefixeCompetence : string = 'PCO';
    PrefixeTypeAction : string = 'RPA';
    lgValeur : integer = 300;

    CODEPRODUIT : string = 'GRC';
    TableLien : String = 'LIENBCONNAISSANCE';
    nb_rattachement : integer = 3;

implementation

uses
 MessagesErreur,PARACTIONS_TOF
 ,Utob,UtilArticle, utofafarticle_select; //mcd 30/08/07

function LanceTheGoodMul(RATTACHEMENT1,RATTACHEMENT2,RATTACHEMENT3,RATTACHEMENT : THVALCOMBOBOX;
  Numchamp : integer) : string;
var chpR : THVALCOMBOBOX;
    retour,stArg : String;
    iNum ,ii: integer;
    vTOBArticle: Tob;
begin
  result := '';
  ChpR:=nil;
  Case Numchamp of
    1 : ChpR :=RATTACHEMENT1;
    2 : ChpR :=RATTACHEMENT2;
    3 : ChpR :=RATTACHEMENT3;
    4 : ChpR :=RATTACHEMENT;
  end;
  if not Assigned (ChpR) then exit;
  iNum:=0;
  stArg:='MULTISELECTION;ACTION=MODIFICATION;RECHERCHEARTICLE';
  if ChpR.value = PrefixeArt then iNum:=NumFicheArt
  else if ChpR.value = PrefixeArtParc then iNum:=NumFicheArtParc
  else if ChpR.value = PrefixeOpe then iNum:=NumFicheOpe
  else if ChpR.value = PrefixeRes then iNum:=NumFicheRes
  else if ChpR.value = PrefixeTiers then begin iNum:=NumFicheTiers; stArg:=';MULTISELECTION'; end
  else if ChpR.value = PrefixeCompetence then iNum:=NumFicheCompetence
  else if ChpR.value = PrefixeTypeAction then iNum:=NumFicheTypeAction
  ;
  if iNum = 0 then exit;
  if (ctxaffaire in v_pgi.PgiCOntexte) and (inum = NumFIcheArt) then
  begin //mcd 30/08/07 GIGA14457 pour OK bon mul article
    VtobArticle:=TOB(StrToInt(AFLanceFiche_Article_MultiSelect( '')));
    if (vTOBArticle <> nil) then
    begin
      for ii := 0 to vTOBArticle.Detail.Count - 1 do
        Retour := Retour+ CodeArticleUnique2(vTOBArticle.detail[ii].GetValue('GA_CODEARTICLE'),'') + ';';
    end;
  end
  else retour:=AGLLanceFiche('RT',Mul_BaseConnaissance[iNum], '','',stArg);
  if Length(retour) > lgValeur then
    begin
    PgiBox(MSG_BaseConnaissance[2],'Base de connaissance');
    iNum:=lgValeur;
    while retour[iNum] <> ';' do iNum := iNum - 1;
    retour:=copy(retour,1,iNum-1);
    end;
  result:=retour;
end;

end.
