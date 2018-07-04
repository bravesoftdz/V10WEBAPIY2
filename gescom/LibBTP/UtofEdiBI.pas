unit UtofEdiBI;

{*****************************************************************
Source  utilisé pour les fiches : RTSUSP_MUL_MAILIN
                                  RTPROS_MUL_MAILIN
                                  RTPROS_MAILIN_FIC
                                  RTPRO_MAILIN_CONT
*****************************************************************}

interface

{
	function ConvertDocFile(
		FileSrce, FileDest: String;
		Initialize, Finalize: TNotifyEvent;
		Funct: TOnFuncEvent;
		EGetVar: TGetVarEvent;
		ESetVar: TSetVarEvent;
		GetList: TOnGetListEvent ): Integer;
}

uses {$IFDEF VER150} variants,{$ENDIF}  Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HMsgBox,UTOF,HQry,HEnt1,StdCtrls,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      mul,Fe_Main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF EAGLCLIENT}
{AFAIREEAGL}
{$ELSE}
Doc_Parser,RT_Parser,
{$ENDIF}
      Hstatus,HDB,M3FP,UTOB, UtilSelection,Paramsoc,UtilPGI,UtilRT,UtilWord,
      AGLUtilOLE,UtilGC,UtilArticle;

Type
    EditionBI = Class (TOF)
    Private
        TobProsp : TOB ;
        stProduitpgi : string;
        StApplication : String;
        Function MailingValidDoc : Boolean;
        procedure ChangeMaquette;
        procedure ChangeDocument;
        procedure Publipostage;
    Public
        procedure OnArgument (Arguments : String ); override ;
        procedure OnLoad ; override ;
        procedure EditionBI(Maquette , DocGenere : string);
        procedure OnGetVar( Sender: TObject; VarName: String; VarIndx: Integer;var Value: variant) ;
     END;

procedure BTPEditionBI(parms:array of variant; nb: integer ) ;
Function BTPMailingValidDoc(parms:array of variant; nb: integer ) : variant ;

implementation

procedure BTPEditionBI(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is EditionBI) then EditionBI(TOTOF).EditionBI(string(Parms[1]),string(Parms[2])) else exit;

end;

Function BTPMailingValidDoc( parms: array of variant; nb: integer ) : variant;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
result:=EditionBI(TOTOF).MailingValidDoc;
end;

procedure EditionBI.OnArgument (Arguments : String );
var    F : TForm;
       x,LaVersionWord : integer;
       NomForme : String;
begin
inherited ;
if (pos('GRF', Arguments) <> 0) then stProduitpgi := 'GRF'
else stProduitpgi := 'GRC';
F := TForm (Ecran);
if stProduitpgi = 'GRF' then
  begin
  if GetParamSoc('SO_RTGESTINFOS003') = True then
      begin
      NomForme :=F.Name;
      F.Name := 'GCFOURNISSEURS';
      MulCreerPagesCL(F, 'Fournisseurs');
      F.Name:=NomForme;
      end;
  end
else if (F.name <> 'RTSUSP_MUL_MAILIN') then MulCreerPagesCL(F, 'Fournisseurs');

if (F.name = 'RTPRO_MAILIN_CONT') and (GetParamSoc('SO_RTGESTINFOS006') = True) then
  begin
  NomForme:=F.Name;
  F.Name:='YYCONTACT';
  MulCreerPagesCL(F, 'Contacts');
  F.Name:=NomForme;
  end;
if (pos('FIC', Arguments) <> 0) then
   begin
   if (F.name = 'RTPRO_MAILIN_CONT') then
      begin
      Ecran.Caption:= 'Export pour Mailing par contact';
      updatecaption(Ecran);
      TFMul(F).Q.Liste:='RTCONTMAILINGFIC';
      SetControlVisible('PDOCUMENT',FALSE) ;
      SetControlVisible('BCREATDOC',FALSE) ;
      SetControlVisible('BMAILINGDOC',FALSE) ;
      Ecran.HelpContext:=111000333 ;
      end;
   end
Else if (pos('REC', Arguments) <> 0) then
   Begin
     StApplication := Arguments;
	 SetControlVisible('PCritere',FALSE);
  	 SetControlVisible('PComplement',FALSE);
	 SetControlVisible('PAvance',FALSE);
 	 SetControlVisible('PSQL',FALSE);
   	 SetControlVisible('PTableLibre',FALSE);
   	 SetControlVisible('PZoneLibre',FALSE);
   	 SetControlVisible('Actions',FALSE);
   	 SetControlVisible('PCONTACT',FALSE);
   	 SetControlVisible('PCONTACTTABLIBRE',FALSE);
	 SetFocusControl('AVEC');
     if stProduitpgi = 'GRF' then
       begin
   	    SetControlText('MAQUETTE',GetParamsoc('SO_RFDIRMAQUETTE')+'\*.doc');
    	SetControlText('DOCGENERE',GetParamsoc('SO_RFDIRSORTIE')+'\DOCUMENTFRS.doc');
	   end
     else
       begin
        SetControlText('MAQUETTE',GetParamsoc('SO_RTDIRMAQUETTE')+'\*.doc');
        SetControlText('DOCGENERE',GetParamsoc('SO_RTDIRSORTIE')+'\MAILING.doc');
       end;
   end
else
    begin
    if stProduitpgi = 'GRF' then
      begin
      SetControlText('MAQUETTE',GetParamsoc('SO_RFDIRMAQUETTE')+'\*.doc');
      SetControlText('DOCGENERE',GetParamsoc('SO_RFDIRSORTIE')+'\DOCUMENTFRS.doc');
      end
    else
      begin
      SetControlText('MAQUETTE',GetParamsoc('SO_RTDIRMAQUETTE')+'\*.doc');
      SetControlText('DOCGENERE',GetParamsoc('SO_RTDIRSORTIE')+'\MAILING.doc');
      end;
    end;
if (F.name = 'RTPRO_MAILIN_CONT') then
   begin
   // LaVersionWord := GetWordVersion;
   LaVersionWord := 9;
   if LaVersionWord<10 then
      begin
      SetControlVisible('ENVOIEMAIL',FALSE) ;
      SetControlVisible('TOBJETMAIL',FALSE) ;
      SetControlVisible('OBJETMAIL',FALSE) ;
      end
   else if GetControlText('ENVOIEMAIL')<>'X' then SetControlEnabled('OBJETMAIL',FALSE) ;
   end;
// Paramétrage des libellés des familles, stat. article et dimensions
if (TFMul(Ecran).name = 'RTLIGN_MUL_MAILIN') or (TFMul(Ecran).name = 'RTLIGN_MAILIN_FIC') or
   (TFMul(Ecran).name = 'RFLIGN_MUL_MAILIN') or (TFMul(Ecran).name = 'RFLIGN_MAILIN_FIC') then
begin
  for x:=1 to 3 do ChangeLibre2('TGA_FAMILLENIV'+InttoStr(x),Ecran);
  ChangeLibre2('TGA_COLLECTION',Ecran);
  MajChampsLibresArticle(TForm(Ecran));
  FiltreComboTypeArticle(THMultiValComboBox(GetControl('TYPEARTICLE')));
end;

end;

procedure EditionBI.OnLoad;
var xx_where,StWhere,StListeActions,StJoint,ListeCombos,Confid : string;
    DateDeb,DateFin : TDateTime;
    i: integer;
begin
inherited ;
if (TFMul(Ecran).name = 'RTSUSP_MUL_MAILIN') then exit;
StWhere := '';
if (GetControlText('SANSSELECT') <> 'X') then
   begin
   DateDeb := StrToDate(GetControlText('DATEACTION'));
   DateFin := StrToDate(GetControlText('DATEACTION_'));
   if GetControlText('SANS') = 'X' then StWhere := 'NOT ';
   if (TFMul(Ecran).name = 'RTPRO_MAILIN_CONT') then
      begin
      StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = RTCONTACTSTIERS.T_AUXILIAIRE';
      StWhere:=StWhere + ' AND RAC_NUMEROCONTACT = RTCONTACTSTIERS.C_NUMEROCONTACT';
      end
   else
       begin
       if stProduitpgi = 'GRF' then StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = RFFOURNISSEURS.T_AUXILIAIRE'
       else StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = RTTIERS.T_AUXILIAIRE';
       end;
   if (GetControlText('TYPEACTION') <> '') and (GetControlText('TYPEACTION') <> '<<Tous>>') then
      begin
      StListeActions:=FindEtReplace(GetControlText('TYPEACTION'),';','","',True);
      StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
      StWhere:=StWhere + ' AND RAC_TYPEACTION in '+StListeActions;
      end;
   if GetControlText('RESPONSABLE') <> '' then
      begin
      ListeCombos := GetControlText('RESPONSABLE');
      if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
      ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
      ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
      StWhere:=StWhere + ' AND RAC_INTERVENANT in '+ListeCombos;
      end;
    for i:=1 to 3 do
        begin
        if (GetControlText('TABLELIBRE'+intToStr(i)) <> '') and (GetControlText('TABLELIBRE'+intToStr(i)) <> '<<Tous>>') then
            begin
            ListeCombos:=FindEtReplace(GetControlText('TABLELIBRE'+intToStr(i)),';','","',True);
            ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
            if stProduitpgi = 'GRF' then StWhere := StWhere + ' AND RAC_TABLELIBREF' +intToStr(i) +' in ' + ListeCombos
            else StWhere := StWhere + ' AND RAC_TABLELIBRE' +intToStr(i) +' in ' + ListeCombos;
            end;
        end;
   if GetControlText('OPERATIONACT') <> '' then
      begin
      StWhere:=StWhere + ' AND RAC_OPERATION = "'+GetControlText('OPERATIONACT')+'"';
      end;
   StWhere := StWhere + ' AND RAC_DATEACTION >= "'+UsDateTime(DateDeb) +'" AND RAC_DATEACTION <= "'+UsDateTime(DateFin)+'"))';
   if (GetControlText('XX_WHERE') = '') then
      SetControlText('XX_WHERE',StWhere)
   else
       begin
       xx_where := GetControlText('XX_WHERE');
       xx_where := xx_where + ' and (' + StWhere + ')';
       SetControlText('XX_WHERE',xx_where) ;
       end;
   end;
{ sélection sur critères lignes/articles }
if (TFMul(Ecran).name = 'RTLIGN_MUL_MAILIN') or (TFMul(Ecran).name = 'RTLIGN_MAILIN_FIC') or
   (TFMul(Ecran).name = 'RFLIGN_MUL_MAILIN') or (TFMul(Ecran).name = 'RFLIGN_MAILIN_FIC') then
    begin
       StWhere := '';
       DateDeb := StrToDate(GetControlText('DATEPIECE'));
       DateFin := StrToDate(GetControlText('DATEPIECE_'));
       if GetControlText('PASDEPIECE') = 'X' then StWhere := 'NOT ';
       if stProduitpgi = 'GRF' then StWhere:=StWhere + 'EXISTS (SELECT GL_NUMERO FROM LIGNE WHERE (GL_TIERS = RFFOURNISSEURS.T_TIERS'
       else StWhere:=StWhere + 'EXISTS (SELECT GL_NUMERO FROM LIGNE WHERE (GL_TIERS = RTTIERS.T_TIERS';
       // NATUREPIECEG,ETABLISSEMENT,DEPOT,REPRESENTANT,ARTICLE,TYPEARTICLE,DOMAINE,
       if (GetControlText('ETABLISSEMENT') <> '') and (GetControlText('ETABLISSEMENT') <> '<<Tous>>') then
          begin
          StListeActions:=FindEtReplace(GetControlText('ETABLISSEMENT'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_ETABLISSEMENT in '+StListeActions;
          end;
       if (GetControlText('NATUREPIECEG') <> '') and (GetControlText('NATUREPIECEG') <> '<<Tous>>') then
          begin
          StListeActions:=FindEtReplace(GetControlText('NATUREPIECEG'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_NATUREPIECEG in '+StListeActions;
          end;
       if (GetControlText('DEPOT') <> '') and (GetControlText('DEPOT') <> '<<Tous>>') then
          begin
          StListeActions:=FindEtReplace(GetControlText('DEPOT'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_DEPOT in '+StListeActions;
          end;
       if (GetControlText('TYPEARTICLE') <> '') and (GetControlText('TYPEARTICLE') <> '<<Tous>>') then
          begin
          StListeActions:=FindEtReplace(GetControlText('TYPEARTICLE'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_TYPEARTICLE in '+StListeActions;
          end;
       if (GetControlText('DOMAINE') <> '') and (GetControlText('DOMAINE') <> '<<Tous>>') then
          begin
          StListeActions:=FindEtReplace(GetControlText('DOMAINE'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_DOMAINE in '+StListeActions;
          end;
       // COLLECTION, FAMILLENIV1,FAMILLENIV2,FAMILLENIV3
       if (GetControlText('COLLECTION') <> '') and (GetControlText('COLLECTION') <> '<<Tous>>') then
          begin
          StListeActions:=FindEtReplace(GetControlText('COLLECTION'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_COLLECTION in '+StListeActions;
          end;
       if (GetControlText('FAMILLENIV1') <> '') and (GetControlText('FAMILLENIV1') <> '<<Tous>>') then
          begin
          StListeActions:=FindEtReplace(GetControlText('FAMILLENIV1'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_FAMILLENIV1 in '+StListeActions;
          end;
       if (GetControlText('FAMILLENIV2') <> '') and (GetControlText('FAMILLENIV2') <> '<<Tous>>') then
          begin
          StListeActions:=FindEtReplace(GetControlText('FAMILLENIV2'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_FAMILLENIV2 in '+StListeActions;
          end;
       if (GetControlText('FAMILLENIV3') <> '') and (GetControlText('FAMILLENIV3') <> '<<Tous>>') then
          begin
          StListeActions:=FindEtReplace(GetControlText('FAMILLENIV3'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_FAMILLENIV3 in '+StListeActions;
          end;
       if GetControlText('ARTICLE') <> '' then
          begin
          StWhere:=StWhere + ' AND GL_CODEARTICLE = "'+GetControlText('ARTICLE')+'"';
          end;
       if (stProduitpgi <> 'GRF') and (GetControlText('REPRESENTANT') <> '') then
          begin
          StWhere:=StWhere + ' AND GL_REPRESENTANT = "'+GetControlText('REPRESENTANT')+'"';
          end;
       for i:=1 to 9 do
          begin
          if (GetControlText('LIBREART'+intToStr(i)) <> '') and (GetControlText('LIBREART'+intToStr(i)) <> '<<Tous>>') then
              begin
              ListeCombos := GetControlText('LIBREART'+intToStr(i));
              if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
              ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
              ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
              StWhere:=StWhere + ' AND GL_LIBREART'+intToStr(i)+' in '+ListeCombos;
              end;
          end;
       if (GetControlText('LIBREARTA') <> '') and (GetControlText('LIBREARTA') <> '<<Tous>>') then
          begin
          ListeCombos := GetControlText('LIBREARTA');
          if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
          ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
          ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
          StWhere:=StWhere + ' AND GL_LIBREARTA in '+ListeCombos;
          end;

       StWhere := StWhere + ' AND GL_DATEPIECE >= "'+UsDateTime(DateDeb) +'" AND GL_DATEPIECE <= "'+UsDateTime(DateFin)+'"))';
       if (GetControlText('XX_WHERE') = '') then
          SetControlText('XX_WHERE',StWhere)
       else
           begin
           xx_where := GetControlText('XX_WHERE');
           xx_where := xx_where + ' and (' + StWhere + ')';
           SetControlText('XX_WHERE',xx_where) ;
           end;
    end;

StWhere := '';
if (GetControlText('COURRIER') <> 'X') then
   begin
   if stProduitpgi = 'GRF' then
     begin
     if GetControlText('FAX') = 'X' then StWhere := '((T_FAX <> "") or (C_FAX <> ""))';
     if GetControlText('NOTFAX') = 'X' then StWhere := '(((T_FAX = "") or (T_FAX is null)) and ((C_FAX = "") or (C_FAX is null))) ';
     end
   else
     begin
     if (TFMul(Ecran).name <> 'RTLIGN_MAILIN_FIC') and (TFMul(Ecran).name <> 'RTLIGN_MUL_MAILIN') then
       begin
       if GetControlText('FAX') = 'X' then StWhere := '((T_FAX <> "") or (C_FAX <> "")) and T_PARTICULIER <> "X"';
       if GetControlText('NOTFAX') = 'X' then StWhere := '(((T_FAX = "") or (T_FAX is null)) and ((C_FAX = "") or (C_FAX is null))) or T_PARTICULIER = "X"';
       end
     else
       begin
       if GetControlText('FAX') = 'X' then StWhere := '(T_FAX <> "") and T_PARTICULIER <> "X"';
       if GetControlText('NOTFAX') = 'X' then StWhere := '((T_FAX = "") or (T_FAX is null)) or T_PARTICULIER = "X"';
       end;
     end;

   if (TFMul(Ecran).name = 'RTPRO_MAILIN_CONT') then
      begin
      if GetControlText('MAIL') = 'X' then StWhere := '(C_RVA <> "")';
      if GetControlText('NOTMAIL') = 'X' then StWhere := '(C_RVA = "")'
      end;
   if (GetControlText('XX_WHERE') = '') then
      SetControlText('XX_WHERE',StWhere)
   else
       begin
       xx_where := GetControlText('XX_WHERE');
       xx_where := xx_where + ' and (' + StWhere + ')';
       SetControlText('XX_WHERE',xx_where) ;
       end;
   end;

  if (V_PGI.MenuCourant = 92 ) then Confid:='CON' else Confid:='CONF';
  if (GetControlText('XX_WHERE') = '') then
      SetControlText('XX_WHERE',RTXXWhereConfident(Confid))
  else
      begin
      xx_where := GetControlText('XX_WHERE');
      xx_where := xx_where + RTXXWhereConfident(Confid);
      SetControlText('XX_WHERE',xx_where) ;
      end;

if (TFMul(Ecran).name = 'RTPROS_MUL_MAILIN') or (TFMul(Ecran).name = 'RTPROS_MAILIN_FIC') or
   (TFMul(Ecran).name = 'RTLIGN_MUL_MAILIN') or (TFMul(Ecran).name = 'RTLIGN_MAILIN_FIC') or
   (TFMul(Ecran).name = 'RFPROS_MUL_MAILIN') or (TFMul(Ecran).name = 'RFPROS_MAILIN_FIC') or
   (TFMul(Ecran).name = 'RFLIGN_MUL_MAILIN') or (TFMul(Ecran).name = 'RFLIGN_MAILIN_FIC') then
    begin
    StJoint := '';
    StJoint := StJoint + 'C_FERME <> "X" and ';
    if (GetCheckBoxState ('PRINCIPAL') <> cbGrayed) then
       StJoint := StJoint + 'C_PRINCIPAL = "' + GetControlText ('PRINCIPAL') + '" and ';
    if (GetCheckBoxState ('PUBLIPOSTAGE') <> cbGrayed) then
       StJoint := StJoint + 'C_PUBLIPOSTAGE = "' + GetControlText ('PUBLIPOSTAGE') + '" and ';
    if (GetControlText('FONCTIONCODEE') <> '') and (GetControlText('FONCTIONCODEE') <> '<<Tous>>') then
       begin
         StJoint := StJoint + 'C_FONCTIONCODEE = "' + GetControlText ('FONCTIONCODEE') + '" and ';
       end;
    For i:=1 to 10 do
        begin
        if i < 10 then
           begin
           if (GetControlText('LIBRECONTACT'+intToStr(i)) <> '') and (GetControlText('LIBRECONTACT'+intToStr(i)) <> '<<Tous>>') then
              begin
              StJoint := StJoint + 'C_LIBRECONTACT' +intToStr(i) +' = "' + GetControlText('LIBRECONTACT'+intToStr(i)) + '" and ';
              end;
           end else
           begin
           if (GetControlText('LIBRECONTACTA') <> '') and (GetControlText('LIBRECONTACTA') <> '<<Tous>>') then
              begin
              StJoint := StJoint + 'C_LIBRECONTACTA' +' = "' + GetControlText('LIBRECONTACTA') + '" and ';
              end;
           end;
        end;
    SetControlText ('XX_JOIN',StJoint);
    end;

if StApplication = 'REC' then SetActiveTabSheet('PACTION');

end;

procedure EditionBI.EditionBI(Maquette , DocGenere : string);
var  F : TFMul;
{$IFDEF EAGLCLIENT}
       L : THGrid;
{$ELSE}
       Ext : String;
       L : THDBGrid;
{$ENDIF}
     i : integer;
     codeprospect,stClauseWhere,StSelect,StSelect2 : string;
     QQ : TQuery ;
     Pages:TPageControl;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
   begin
   PGIInfo('Aucun élément sélectionné','');
   exit;
   end;
if PGIAsk('Confirmez-vous le traitement ?','')<>mrYes then exit ;

L:= F.FListe;
Pages:=F.Pages;

StSelect := 'Select T_TIERS,T_AUXILIAIRE,T_NATUREAUXI,T_LIBELLE,T_ADRESSE1,T_ADRESSE2,T_ADRESSE3';
StSelect := StSelect + ',T_VILLE,T_CODEPOSTAL,T_TELEPHONE,T_SIRET,T_FAX,T_JURIDIQUE,T_PAYS';
StSelect := StSelect + ',T_TELEX,T_ZONECOM,T_REPRESENTANT,T_APE,T_PRENOM,T_SOCIETEGROUPE,T_PRESCRIPTEUR';
StSelect := StSelect + ',T_PARTICULIER,C_NOM,C_PRENOM,C_CIVILITE,C_FONCTION,C_SERVICE';
StSelect := StSelect + ',RPR_RPRLIBBOOL0,RPR_RPRLIBBOOL1,RPR_RPRLIBBOOL2,RPR_RPRLIBBOOL3';
StSelect := StSelect + ',RPR_RPRLIBBOOL4,RPR_RPRLIBBOOL5,RPR_RPRLIBBOOL6,RPR_RPRLIBBOOL7';
StSelect := StSelect + ',RPR_RPRLIBBOOL8,RPR_RPRLIBBOOL9';
StSelect := StSelect + ',RPR_RPRLIBTABLE0,RPR_RPRLIBTABLE1,RPR_RPRLIBTABLE2,RPR_RPRLIBTABLE3,RPR_RPRLIBTABLE4';
StSelect := StSelect + ',RPR_RPRLIBTABLE5,RPR_RPRLIBTABLE6,RPR_RPRLIBTABLE7,RPR_RPRLIBTABLE8,RPR_RPRLIBTABLE9';
StSelect := StSelect + ',RPR_RPRLIBTABLE10,RPR_RPRLIBTABLE11,RPR_RPRLIBTABLE12,RPR_RPRLIBTABLE13,RPR_RPRLIBTABLE14';
StSelect := StSelect + ',RPR_RPRLIBTABLE15,RPR_RPRLIBTABLE16,RPR_RPRLIBTABLE17,RPR_RPRLIBTABLE18,RPR_RPRLIBTABLE19';
StSelect := StSelect + ',RPR_RPRLIBTABLE20,RPR_RPRLIBTABLE21,RPR_RPRLIBTABLE22,RPR_RPRLIBTABLE23,RPR_RPRLIBTABLE24';
StSelect := StSelect + ',RPR_RPRLIBTABLE25';
StSelect := StSelect + ',RPR_RPRLIBVAL0,RPR_RPRLIBVAL1,RPR_RPRLIBVAL2,RPR_RPRLIBVAL3,RPR_RPRLIBVAL4';
StSelect := StSelect + ',RPR_RPRLIBVAL5,RPR_RPRLIBVAL6,RPR_RPRLIBVAL7,RPR_RPRLIBVAL8,RPR_RPRLIBVAL9';
StSelect := StSelect + ',RPR_RPRLIBDATE0,RPR_RPRLIBDATE1,RPR_RPRLIBDATE2,RPR_RPRLIBDATE3,RPR_RPRLIBDATE4';
StSelect := StSelect + ',RPR_RPRLIBDATE5,RPR_RPRLIBDATE6,RPR_RPRLIBDATE7,RPR_RPRLIBDATE8,RPR_RPRLIBDATE9';
StSelect := StSelect + ',RPR_RPRLIBTEXTE0,RPR_RPRLIBTEXTE1,RPR_RPRLIBTEXTE2,RPR_RPRLIBTEXTE3,RPR_RPRLIBTEXTE4';
StSelect := StSelect + ',RPR_RPRLIBTEXTE5,RPR_RPRLIBTEXTE6,RPR_RPRLIBTEXTE7,RPR_RPRLIBTEXTE8,RPR_RPRLIBTEXTE9';
StSelect := StSelect + ',YTC_TABLELIBRETIERS1,YTC_TABLELIBRETIERS2,YTC_TABLELIBRETIERS3';
StSelect := StSelect + ',YTC_TABLELIBRETIERS4,YTC_TABLELIBRETIERS5,YTC_TABLELIBRETIERS6';
StSelect := StSelect + ',YTC_TABLELIBRETIERS7,YTC_TABLELIBRETIERS8,YTC_TABLELIBRETIERS9';
StSelect := StSelect + ',YTC_TABLELIBRETIERSA,YTC_BOOLLIBRE1,YTC_BOOLLIBRE2,YTC_BOOLLIBRE3';
StSelect := StSelect + ',YTC_VALLIBRE1,YTC_VALLIBRE2,YTC_VALLIBRE3,YTC_DATELIBRE1,YTC_DATELIBRE2,YTC_DATELIBRE3';
StSelect := StSelect + ',YTC_TEXTELIBRE1,YTC_TEXTELIBRE2,YTC_TEXTELIBRE3';
StSelect := StSelect + ' FROM RTTIERS';
StSelect := StSelect + ' left join CONTACT on C_AUXILIAIRE=T_AUXILIAIRE and C_TYPECONTACT = "T" and C_PRINCIPAL = "X" ';


TobProsp:=Tob.create('_Les Prospects',nil,-1);

if L.AllSelected then
   begin
   stClauseWhere:=RecupWhereCritere(Pages);
   StSelect := StSelect + stClauseWhere;
   QQ:=OpenSQL(StSelect,TRUE) ;
   if Not QQ.EOF then TobProsp.LoadDetailDB('','','',QQ,False,True);
   Ferme(QQ) ;
   L.AllSelected:=False;
   end else
   begin
   InitMove(L.NbSelected,'');
   for i:=0 to L.NbSelected-1 do
      begin
      MoveCur(False);
      L.GotoLeBookmark(i);
      Tob.create('',TobProsp,-1);
      codeprospect:=TFmul(Ecran).Q.FindField('T_TIERS').asstring ;
      StSelect2 := StSelect + ' WHERE T_TIERS ="'+codeprospect+'"';
      QQ:=OpenSQL(StSelect2,TRUE) ;
      if Not QQ.EOF then TobProsp.detail[i].selectDB('',QQ,false);
      Ferme(QQ) ;
      end;
   L.ClearSelected;
   end;
FiniMove;

{$IFDEF EAGLCLIENT}
{AFAIREEAGL}
{$ELSE}
Ext := ExtractFileExt(Maquette);
if Ext <> '' then
   begin
   if Ext = '.doc' then ConvertDocFile(Maquette,DocGenere,Nil,Nil,Nil,OnGetVar,Nil,Nil)
   else ConvertRTFFile(Maquette,DocGenere,Nil,Nil,Nil,OnGetVar,Nil,Nil) ;
   end;
{$ENDIF}
TobProsp.free;
TobProsp:=Nil;
end ;


procedure EditionBI.OnGetVar( Sender: TObject; VarName: String; VarIndx: Integer;var Value: variant) ;
var TTOB : tob ;
BEGIN
if VarName='NBPROSPECT' then
   begin
      Value:=TobProsp.Detail.Count;
      exit;
   end;
if (Pos('T_',VarName)<= 0) and (Pos('RPR_',VarName)<=0) and (Pos('YTC_',VarName)<=0) and (Pos('C_',VarName)<=0) then exit;
TTob:=TobProsp.Detail[VarIndx-1];
if TTOB.GetValue(VarName) = NULL then
   begin
   if ((Pos('RPRLIBVAL',VarName)>0) or (Pos('VALLIBRE',VarName)>0)) then TTOB.PutValue(VarName,0)
   else if ((Pos('RPRLIBDATE',VarName)>0) or (Pos('DATELIBRE',VarName)>0)) then TTOB.PutValue(VarName,IDate1900)
   else TTOB.PutValue(VarName,'');
   end;
if Pos('BLOCNOTE',VarName)>0
      then  begin BlobToFile(VarName,TTob.GetValue(VarName)); Value:=''; end else
            if ChampToType( VarName)= 'COMBO' then Value:=RechDom(ChampToTT( VarName),TTob.GetValue(VarName),false)
               else Value:=TTOB.GetValue(VarName) ;
END ;

Function EditionBI.MailingValidDoc : Boolean;
var stDocWord , stMaquette: string;
begin
Result:=False;
stDocWord:=getControlText('DOCGENERE');
stMaquette:=getControlText('MAQUETTE');
if (stMaquette='') or (pos('\*.',stMaquette)<>0) then begin PGIBox('Vous devez choisir une maquette', ecran.caption); exit; end;
if not FileExists(stMaquette) then begin PGIBox('La maquette n''existe pas', ecran.caption); exit; end;
if (stDocWord='') then begin PGIBox('Vous devez mentionner un document', ecran.caption); exit; end;
if FileExists(stDocWord) then
   if PGIAsk('Ce fichier existe déjà, voulez-vous l''écraser ?',ecran.caption)=mryes then DeleteFile(stDocWord ) else exit;
if (TFMul(Ecran).name = 'RTPRO_MAILIN_CONT') then
   begin
   if GetControlText('ENVOIEMAIL') = 'X' then
     begin
     if GetControlText('OBJETMAIL') = '' then begin PGIBox('Vous devez renseigner l''objet du message', ecran.caption); exit; end;
    //   StSql := TFMul(Ecran).Q.sql.Text;
    //   if (pos('C_RVA',StSql) = 0) or (pos('C_RVA',StSql) > pos('FROM',StSql)) then begin PGIBox('Vous devez ajouter le champ E-mail dans la liste', ecran.caption); exit; end;
     if GetControlText('MAIL') <> 'X' then begin PGIBox('Vous devez sélectionner les contacts avec e-mail (Onglet compléments)', ecran.caption); exit; end;
//     if GetControlText('C_EMAILING') <> 'X' then begin PGIBox('Vous devez sélectionner les contacts avec e-mail opt-in', ecran.caption); exit; end;
     end;
   end;
Result:=True;
end;


procedure EditionBI.ChangeMaquette;
var stMaquette: string;
    OkOk : Boolean ;
begin
OkOk:=False;
stMaquette:=getControlText('MAQUETTE');
if (stMaquette<>'') and (pos('\*.',stMaquette)=0) and  FileExists(stMaquette) then OkOk:=true;
SetControlEnabled('BMAQUETTE',  OkOk) ;
//SetControlEnabled('BOUVRIR',  OkOk) ;
SetControlEnabled('BMAILINGDOC',  OkOk) ;
end;

procedure EditionBI.ChangeDocument;
var stDocument: string;
    OkOk : boolean ;
begin
OkOk:=False;
stDocument:=getControlText('DOCGENERE');
if (stDocument<>'') and (pos('\*.',stDocument)=0) and  FileExists(stDocument) then OkOk:=true;
SetControlEnabled('BDOCUMENT',OkOk) ;
end;

procedure EditionBI.Publipostage;
var StTitre,stDocument,stMaquette,ArgEmail: string;
    DS : THquery;
begin
stMaquette:=getControlText('MAQUETTE');
stDocument:=getControlText('DOCGENERE');

if (TFMul(Ecran).name = 'RTPRO_MAILIN_CONT') then
   begin
   if GetControlText('ENVOIEMAIL') = 'X' then
      begin
      {$IFDEF EAGLCLIENT}
      if TFMul(Ecran).Fetchlestous then
      begin
      {$ENDIF}
      ArgEmail := 'EMAIL;-;ADRESSEEMAIL;';
      ArgEmail := ArgEmail + GetControlText('OBJETMAIL');
      DS:=THquery(TFMul(Ecran).Q) ;
      StTitre :=  TFMul(Ecran).Q.titres ;// GetTitreGrid (TFMul(Ecran).FListe);
      LancePublipostage(ArgEmail,stMaquette,stDocument,Nil,StTitre,DS,True);
  {$IFDEF EAGLCLIENT}
      end;
  {$ENDIF}
      exit;
      end;
   end;

{$IFDEF EAGLCLIENT}
if TFMul(Ecran).Fetchlestous then
   begin
{$ENDIF} //EAGLCLIENT
   DS:=THquery(TFMul(Ecran).Q) ;
   StTitre :=  TFMul(Ecran).Q.titres ;// GetTitreGrid (TFMul(Ecran).FListe);
   LancePublipostage('FILE',stMaquette,stDocument,Nil,StTitre,DS,True)
{$IFDEF EAGLCLIENT}
   end;
{$ENDIF}
end;


Procedure BTPMailing_ChangeMaquette( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFmul) then ToTof:=TFmul(F).LaTof else exit;
  EditionBI(totof).ChangeMaquette;
end;
Procedure BTPMailing_ChangeDocument( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFmul) then ToTof:=TFmul(F).LaTof else exit;
  EditionBI(totof).ChangeDocument;
end;

Procedure BTPPublipostage( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFmul) then ToTof:=TFmul(F).LaTof else exit;
  EditionBI(totof).Publipostage;
end;

Function BTPGetTitreGrid( parms: array of variant; nb: integer ) : variant ;
var  F : TForm ;
begin
  Result:='';
  F:=TForm(Longint(Parms[0])) ;
//  if (F is TFmul) then result:=getTitreGrid(TFmul(F).Fliste) else exit;
  if (F is TFmul) then result:=TFMul(F).Q.titres else exit;
end;

Initialization
registerclasses([EditionBI]);
RegisterAglProc('EditionBI',TRUE,2,BTPEditionBI);
RegisterAglFunc('MailingValidDoc', TRUE , 0, BTPMailingValidDoc);
RegisterAglProc( 'Mailing_ChangeMaquette', TRUE ,0 , BTPMailing_ChangeMaquette);
RegisterAglProc( 'Mailing_ChangeDocument', TRUE , 0, BTPMailing_ChangeDocument);
RegisterAglFunc( 'GetTitreGrid', TRUE ,0 , BTPGetTitreGrid);
RegisterAglProc( 'Publipostage', TRUE ,0 , BTPPublipostage);
end.
