unit UtofMailingAff;

interface
uses  Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,UTOF,HEnt1,StdCtrls,menus,Hstatus,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      mul,Fe_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
      UtilSelection,Paramsoc,UtilRT,UTOB,UtilPGI,
      utofAfBaseCodeAffaire,
{$Ifdef GIGI}
      Dicobtp,
{$ENDIf}
      HMsgBox,
{$ifndef BUREAU}
      UTOFDOCEXTERNE,
{$endif}
      HQry,HTB97,UtilWord,UtilGC,EntGC,UtilGA,UtofMailing
{$IF Defined(CRM)}
      ,UtofRTSaisieInfosEsker,EskerInterface,Web,AglInit
{$IFEND}
      ;

Type
    TOF_MailingAff = Class (TOF_AFBASECODEAFFAIRE)
    Private
        stMaquetteValide : string;
        stNatureDoc : string;
{$IF Defined(CRM)}
        TobProsp : TOB ;
        TobInfos : Tob;
        TypeEnvoiEsker : string;
        TypeTraitement : string;
        TobAdr : Tob;
{$IFEND}
{$ifndef BUREAU}
        procedure OpenDoc_OnClick(Sender: TObject);
        procedure MAQUETTE_OnElipsisClick(Sender: TObject);
        procedure ChangeMaquette(Sender: TObject);
        procedure BMAILINGDOC_OnClick(Sender: TObject);
{$endif}
        procedure BEMAILING_OnClick(Sender: TObject);
        Function MailingValidDoc : Boolean;
{$IF Defined(CRM)}
        procedure BESKERCOURRIER_OnClick(Sender: TObject);
        procedure BESKEREMAIL_OnClick(Sender: TObject);
        procedure BESKERFAX_OnClick(Sender: TObject);
        procedure BESKERSUIVI_OnClick(Sender: TObject);
        function ChpsObligExiste : Boolean;
        procedure PublipostageCallBack(Index: integer; OutputFileName: hstring; var Cancel: boolean);
        procedure CreatTobInfosMul ;
{$IFEND}
    Protected
    {$IFDEF AFFAIRE}
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override;
        procedure bSelectAff1Click(Sender: TObject);     override ;
    {$ENDIF AFFAIRE}
    Public
        procedure OnArgument (Arguments : String ); override ;
        procedure OnLoad ; override ;
        procedure OnClose ; override ;
        procedure BVALIDSELClick(Sender: TObject);
        procedure CodeArticleRech(Sendre: TOBject);
        procedure constitueXXWHEREPARC;
        procedure ScreenKeYDown(Sender: TObject; var Key: Word;Shift: TShiftState);
     END;

Procedure RTLanceFiche_MailingAffContactExp;
Procedure RTLanceFiche_MailingAffContact;
Procedure RTLanceFiche_GenActionsAffContact;

const
	// libellés des messages
	TexteMessage: array[1..3] of string 	= (
          {1}        'Il manque les champs E-Mail (contact et tiers) dans la liste'
          {2}        ,'Il manque les champs Fax (contact et tiers) dans la liste'
          {3}        ,'Envoi n° %s'
          );

implementation

uses ed_tools,AglInitGC;

Procedure RTLanceFiche_MailingAffContactExp;
begin
  AGLLanceFiche('RT','RTAFF_MAILIN_CONT','','','FIC') ;
end;

Procedure RTLanceFiche_MailingAffContact;
begin
  AGLLanceFiche('RT','RTAFF_MAILIN_CONT','','','MAI') ;
end;

Procedure RTLanceFiche_GenActionsAffContact;
begin
  AGLLanceFiche('RT','RTGENERE_ACTIONS','','','AFF');
end;

procedure TOF_MailingAff.OnArgument (Arguments : String );
var Critere,ChampMul,ValMul : string;
    x,iCol : integer;
    F : TForm;
begin
inherited ;
F := TForm (Ecran);
MulCreerPagesCL(F,'NOMFIC=GCTIERS');
if (GetParamSocSecur('SO_RTGESTINFOS006',False) = True) then
    MulCreerPagesCL(F,'NOMFIC=YYCONTACT');

if (arguments='MAI') then
  begin
//    SetControlText('MAQUETTE',GetParamsocSecur('SO_RTDIRMAQUETTE','C:\PGI00\STD\MAQUETTE')+'\*.doc');
    SetControlText('DOCGENERE',GetParamsocSecur('SO_RTDIRSORTIE','')+'\MAILING.doc');
    SetControlEnabled('BMAILINGDOC',FALSE) ;
    SetControlEnabled('BEMAILING',FALSE) ;
    stNatureDoc := 'MLC';
  end
else if (arguments='FIC') then
  begin
    Ecran.Caption:= TraduireMemoire('Export pour Mailing par contact des affaires');
    updatecaption(Ecran);
//    TFMul(F).Q.Liste:='RTAFFMAILINGFIC';
    TFMul(F).SetDBListe('RTAFFMAILINGFIC');
    SetControlVisible('PDOCUMENT',FALSE) ;
    SetControlVisible('BCREATDOC',FALSE) ;
    SetControlVisible('BMAILINGDOC',FALSE) ;
    SetControlVisible('BEMAILING',FALSE) ;
    Ecran.HelpContext:=111000333 ;
  end
else
  begin  //arguments : OPERATION='+operation,'','CODEOPER='+operation+';ACTGEN='+act
  Repeat
    Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
    if Critere <> '' then
    begin
      x := pos('=',Critere);
      if x <> 0 then
      begin
        ChampMul:=copy(Critere,1,x-1);
        ValMul:=copy(Critere,x+1,length(Critere));
        if ChampMul='CODEOPER' then
        begin
          setcontroltext('OPGENACT',ValMul);
        end;
        if ChampMul='ACTGEN' then
        begin
          THNumedit(getcontrol('NUMGENACT')).value:=strtoint(ValMul);
        end;
      end;
    end;
  until  Critere='';
  SetControlVisible('BOUVRIR',True) ;
  SetControlVisible('BSelectAll',True) ;
{$IFDEF EAGLCLIENT}
  SetControlProperty('FLISTE','Multiselect',True) ;
{$ELSE}
  SetControlProperty('FLISTE','Multiselection',True) ;
{$ENDIF}
  SetControlVisible('PDOCUMENT',False) ;
  SetControlVisible('BCREATDOC',False) ;
  SetControlVisible('BMAILINGDOC',False) ;
  SetControlVisible('BEMAILING',False) ;
  if (getcontroltext('OPGENACT') = '') then SetControlVisible('ACTIONS',False) ;
  Ecran.Caption:= TraduireMemoire('Génération des actions par contact des affaires');
  updatecaption(Ecran);
//  TFMul(F).Q.Liste:='RTAFFGENACTION';
  TFMul(F).SetDBListe('RTAFFGENACTION');
  Ecran.HelpContext:=111000346 ;
  SetControlVisible('OPERATION',False) ;
  SetControlVisible('TOPERATION_',False) ;
  SetControlVisible('CHOIXENVOI',False) ;
  SetControlVisible('COURRIER',False) ;
  SetControlVisible('FAX',False) ;
  SetControlVisible('NOTFAX',False) ;
  end;
  for iCol:=1 to 9 do ChangeLibre2('TAFF_LIBREAFF'+InttoStr(iCol),Ecran);
  ChangeLibre2('TAFF_LIBREAFFA',Ecran);
  for iCol:=1 to 3 do ChangeLibre2('TAFF_VALLIBRE'+InttoStr(iCol),Ecran);
  for iCol:=1 to 3 do ChangeLibre2('TAFF_DATELIBRE'+InttoStr(iCol),Ecran);
  for iCol:=1 to 3 do ChangeBoolLibre('AFF_BOOLLIBRE'+InttoStr(iCol),Ecran);
  for iCol:=1 to 3 do SetControlCaption('BOOLLIBRE'+InttoStr(iCol),TCheckBox(GetControl('AFF_BOOLLIBRE'+InttoStr(iCol))).caption);
  for iCol:=1 to 3 do
      if copy(TCheckBox(GetControl('AFF_BOOLLIBRE'+InttoStr(iCol))).caption,1,2)='.-' then
         SetControlVisible('BOOLLIBRE'+InttoStr(iCol),false);
  for iCol:=1 to 3 do ChangeLibre2('TAFF_RESSOURCE'+InttoStr(iCol),Ecran);

{  GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'VALLIBRE', 3, '_');
  GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'DATELIBRE', 3, '_');
  GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'BOOLLIBRE', 3, '_');
  GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'RESSOURCE', 3, '_');
}
if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
  if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
  else begin
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
    if not (ctxscot in V_PGI.PGICOntexte) then
       begin
       SetControlVisible ('T_MOISCLOTURE',false);
       SetControlVisible ('T_MOISCLOTURE_',false);
       SetControlVisible ('TT_MOISCLOTURE',false);
       SetControlVisible ('TT_MOISCLOTURE_',false);
       end;
    end;
  end;
{$Ifdef GIGI}
 if (GetControl('T_REPRESENTANT') <> nil) then  SetControlVisible('T_REPRESENTANT',false);
 if (GetControl('TT_REPRESENTANT') <> nil) then  SetControlVisible('TT_REPRESENTANT',false);
 if (GetControl('T_ZONECOM') <> nil) then  SetControlVisible('T_ZONECOM',false);
 if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
 SetControlText('T_NatureAuxi','');    //on efface les valeurs CLI et PO, car NCP en plus
 SetControlProperty ('T_NATUREAUXI', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI_', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI_', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI_', 'Plus', VH_GC.AfNatTiersGRCGI);
 SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
  //MCD 04/07/2005
 if (F.name = 'RTAFF_MAILIN_CONT') then
  begin //on peut avoir des info du DP
  //mcd 20/09/2005 agl 148 TfMul(Ecran).dbliste := 'AFAFFMAILINGCON';
  //mcd 20/09/2005 agl 148 if TfMul(Ecran).Q <> NIL then TfMul(Ecran).Q.Liste  := 'AFAFFMAILINGCON';
  if Ecran.HelpContext=111000346 then
   begin   // liste # si lancmeent opération ou mailing
   TfMul(Ecran).Setdbliste( 'AFAFFGENACTION');
   Ecran.Caption:=TraduitGa( 'Génération des actions par contact des affaires');
   updatecaption(Ecran);
   end
  else  TfMul(Ecran).Setdbliste( 'AFAFFMAILINGCON');
  MulCreerPagesDP(F);
  end;
{$endif}
{$ifndef BUREAU}
  if Assigned(GetControl('OpenDoc')) then
    TMenuItem(GetControl('OpenDoc')).OnClick := OpenDoc_OnClick;
  if Assigned(GetControl('MAQUETTE')) then
    begin
    THEDIT(GetControl('MAQUETTE')).OnElipsisClick := MAQUETTE_OnElipsisClick;
    THEDIT(GetControl('MAQUETTE')).OnChange := ChangeMaquette;
  if Assigned(GetControl('BMAILINGDOC')) then
    TToolBarButton97(GetControl('BMAILINGDOC')).OnClick := BMAILINGDOC_OnClick;
    end;
{$else}    //dans le bureau, on envuet pas les maquetts. sinon pb avec affaireOle
  SetControlVisible ('MAQUETTE',false);
  TMenuItem(GetControl('OpenDoc')).visible := false;
{$endif}
  if Assigned(GetControl('BEMAILING')) then
    TToolBarButton97(GetControl('BEMAILING')).OnClick := BEMAILING_OnClick;
{$IF Defined(CRM)}
  if (arguments='MAI') then
    begin
    if GetParamSocSecur ('SO_RTGESTIONFLYDOC',False) = True then
      begin
      SetControlVisible ('BESKERSUIVI',True);
      SetControlVisible ('BESKERCOURRIER',True);
      SetControlVisible ('BESKEREMAIL',True);
      SetControlVisible ('BESKERFAX',True);
      SetControlEnabled('BESKERCOURRIER',False) ;
      SetControlEnabled('BESKEREMAIL',False) ;
      SetControlEnabled('BESKERFAX',False) ;
      end;
    if Assigned(GetControl('BESKERCOURRIER')) then
       TToolbarButton97(GetControl('BESKERCOURRIER')).OnClick := BESKERCOURRIER_OnClick;
    if Assigned(GetControl('BESKEREMAIL')) then
       TToolbarButton97(GetControl('BESKEREMAIL')).OnClick := BESKEREMAIL_OnClick;
    if Assigned(GetControl('BESKERFAX')) then
       TToolbarButton97(GetControl('BESKERFAX')).OnClick := BESKERFAX_OnClick;
    if Assigned(GetControl('BESKERSUIVI')) then
       TToolbarButton97(GetControl('BESKERSUIVI')).OnClick := BESKERSUIVI_OnClick;
    end;
{$IFEND}

  if TToolbarButton97 (GetControl('BVALIDSEL')) <> nil then
  begin
    TToolbarButton97 (GetControl('BVALIDSEL')).OnClick := BVALIDSELClick;
    TToolbarButton97 (GetControl('BVALIDSEL')).Left := THValComboBox (GetControl('BCherche')).Left;
  end;
  if not JaiLeDroitTag(323210) then
  begin
    TTabSheet (GetControl('TBPARCSAV')).TabVisible := false;
  end;
  if ThEdit(GetControl('CODEARTICLE')) <> nil then
  begin
    ThEdit(GetControl('CODEARTICLE')).OnElipsisClick := CodeArticleRech;
  end;
  Ecran.OnKeyDown := ScreenKeYDown;


end;

procedure TOF_MailingAff.OnLoad;
var xx_where,StWhere,StListeActions,ListeCombos: string;
    DateDeb,DateFin : TDateTime;
    ValDeb,ValFin : double;
    i : integer;
begin
inherited ;
// PL 30/08/07 : c'était dans le script, c'est plus propre dans la tof...et en plus ça ne marchait pas en GI
if (GetCheckBoxState('EXISTE') = cbGrayed) then
   SetControlText('XX_WHERE','')
 else
   begin
   if (GetCheckBoxState('EXISTE') <> cbChecked) then
      StWhere := 'not '
   else
      StWhere := '';

 {$ifdef GIGI}
   StWhere := StWhere + ' exists (select RAC_NUMACTGEN from ACTIONS  where ((RAC_AUXILIAIRE=AFDPCONTTIERS.T_AUXILIAIRE) AND RAC_NUMEROCONTACT=AFDPCONTTIERS.C_NUMEROCONTACT AND RAC_OPERATION="'
        + GetControlText('OPERATION') + '" AND (RAC_NUMACTGEN='
        + GetControlText('NUMACTGEN') + ')))';
  {$else}
   StWhere := StWhere + ' exists (select RAC_NUMACTGEN from ACTIONS  where ((RAC_AUXILIAIRE=RTCONTACTSTIERS.T_AUXILIAIRE) AND RAC_NUMEROCONTACT=RTCONTACTSTIERS.C_NUMEROCONTACT AND RAC_OPERATION="'
        + GetControlText('OPERATION') + '" AND (RAC_NUMACTGEN='
        + GetControlText('NUMACTGEN') + ')))';
 {$ENDIF GIGI}

   SetControlText('XX_WHERE', StWhere);
   end;

SetControlProperty('NUMACTGEN','Plus',GetControlText ('OPERATION')) ;
SetControlProperty('NUMACTGEN','DataType','RTACTGEN') ;
StWhere := '';
if (GetControlText('SANSSELECT') <> 'X') then
   begin
   DateDeb := StrToDate(GetControlText('DATEACTION'));
   DateFin := StrToDate(GetControlText('DATEACTION_'));
   if (GetControlText('SANS') = 'X') then StWhere := 'NOT ';
   //StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = T_AUXILIAIRE';
   //StWhere:=StWhere + ' AND RAC_NUMEROCONTACT = C_NUMEROCONTACT';
 {$ifdef GIGI}  //MCD 04/07/2005
   StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = AFDPCONTTIERS.T_AUXILIAIRE';
   StWhere:=StWhere + ' AND RAC_NUMEROCONTACT = AFDPCONTTIERS.C_NUMEROCONTACT';
  {$else}
   StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = RTCONTACTSTIERS.T_AUXILIAIRE';
   StWhere:=StWhere + ' AND RAC_NUMEROCONTACT = RTCONTACTSTIERS.C_NUMEROCONTACT';
 {$ENDIF GIGI}

   if GetControlText('TYPEACTION') <> '' then
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
        if (GetControlText('TABLELIBRE'+intToStr(i)) <> '') and (GetControlText('TABLELIBRE'+intToStr(i)) <> TraduireMemoire('<<Tous>>')) then
            begin
            ListeCombos:=FindEtReplace(GetControlText('TABLELIBRE'+intToStr(i)),';','","',True);
            ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
            StWhere := StWhere + ' AND RAC_TABLELIBRE' +intToStr(i) +' in ' + ListeCombos;
            end;
        end;
   if GetControlText('OPERATIONACT') <> '' then
      begin
      StWhere:=StWhere + ' AND RAC_OPERATION = "'+GetControlText('OPERATIONACT')+'"';
      end;
   if (GetControlText('ETATACTION') <> '') and (GetControlText('ETATACTION') <> TraduireMemoire('<<Tous>>')) then
      begin
      ListeCombos:=FindEtReplace(GetControlText('ETATACTION'),';','","',True);
      ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
      StWhere:=StWhere + ' AND RAC_ETATACTION in '+ListeCombos;
      end;
   if (GetControlText('NIVIMP') <> '') and (GetControlText('NIVIMP') <> TraduireMemoire('<<Tous>>')) then
      begin
      ListeCombos:=FindEtReplace(GetControlText('NIVIMP'),';','","',True);
      ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
      StWhere:=StWhere + ' AND RAC_NIVIMP in '+ListeCombos;
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
{ mng 12/01/2004 : sélection sur critères affaires/articles }
   StWhere := '';
   DateDeb := StrToDate(GetControlText('DATEDEBUT'));
   DateFin := StrToDate(GetControlText('DATEDEBUT_'));
 {$ifdef GIGI}   //MCD 04/07/2005
   StWhere:=StWhere + 'EXISTS (SELECT AFF_AFFAIRE FROM AFFAIRE WHERE (AFF_TIERS = AFDPCONTTIERS.T_TIERS';
   StWhere:=StWhere + ' AND AFF_NUMEROCONTACT = AFDPCONTTIERS.C_NUMEROCONTACT';
  {$else}
   StWhere:=StWhere + 'EXISTS (SELECT AFF_AFFAIRE FROM AFFAIRE WHERE (AFF_TIERS = RTCONTACTSTIERS.T_TIERS';
   StWhere:=StWhere + ' AND AFF_NUMEROCONTACT = RTCONTACTSTIERS.C_NUMEROCONTACT';
 {$ENDIF GIGI}

   if GetControlText('AFFAIRE1') <> '' then
      StWhere:=StWhere + ' AND AFF_AFFAIRE1 LIKE "'+GetControlText('AFFAIRE1')+'%"';
   if GetControlText('AFFAIRE2') <> '' then
      StWhere:=StWhere + ' AND AFF_AFFAIRE2 LIKE "'+GetControlText('AFFAIRE2')+'%"';
   if GetControlText('AFFAIRE3') <> '' then
      StWhere:=StWhere + ' AND AFF_AFFAIRE3 LIKE "'+GetControlText('AFFAIRE3')+'%"';

   if GetControlText('AVENANT') <> '' then
      StWhere:=StWhere + ' AND AFF_AVENANT LIKE "'+GetControlText('AVENANT')+'%"';

   if GetControlText('RESPONSABLEAFF') <> '' then
      StWhere:=StWhere + ' AND AFF_RESPONSABLE = "'+GetControlText('RESPONSABLEAFF')+'"';

   if GetControlText('STATUTAFFAIRE') <> '' then
      StWhere:=StWhere + ' AND AFF_STATUTAFFAIRE = "'+GetControlText('STATUTAFFAIRE')+'"';

   if GetControlText('ETATAFFAIRE') <> '' then
      StWhere:=StWhere + ' AND AFF_ETATAFFAIRE = "'+GetControlText('ETATAFFAIRE')+'"';

   if GetControlText('RESSOURCE1') <> '' then
      StWhere:=StWhere + ' AND AFF_RESSOURCE1 = "'+GetControlText('RESSOURCE1')+'"';
   if GetControlText('RESSOURCE2') <> '' then
      StWhere:=StWhere + ' AND AFF_RESSOURCE2 = "'+GetControlText('RESSOURCE2')+'"';
   if GetControlText('RESSOURCE3') <> '' then
      StWhere:=StWhere + ' AND AFF_RESSOURCE3 = "'+GetControlText('RESSOURCE3')+'"';

   for i:=1 to 9 do
      begin
      if (GetControlText('LIBREAFF'+intToStr(i)) <> '') and (GetControlText('LIBREAFF'+intToStr(i)) <> TraduireMemoire('<<Tous>>')) then
          begin
          ListeCombos := GetControlText('LIBREAFF'+intToStr(i));
          if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
          ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
          ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
          StWhere:=StWhere + ' AND AFF_LIBREAFF'+intToStr(i)+' in '+ListeCombos;
          end;
      end;
   if (GetControlText('LIBREAFFA') <> '') and (GetControlText('LIBREAFFA') <> TraduireMemoire('<<Tous>>')) then
      begin
      ListeCombos := GetControlText('LIBREAFFA');
      if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
      ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
      ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
      StWhere:=StWhere + ' AND AFF_LIBREAFFA in '+ListeCombos;
      end;

   if (GetControlText('DATEDEBUT') <> DateTimeToStr(iDate1900))
      and (GetControlText('DATEDEBUT')+'_' <> DateTimeToStr(iDate2099)) then
      StWhere := StWhere + ' AND AFF_DATEDEBUT >= "'+UsDateTime(DateDeb) +'" AND AFF_DATEDEBUT <= "'+UsDateTime(DateFin)+'"';

   if (GetControlText('DATEFIN') <> DateTimeToStr(iDate1900))
      and (GetControlText('DATEFIN')+'_' <> DateTimeToStr(iDate2099)) then
       begin
       DateDeb := StrToDate(GetControlText('DATEFIN'));
       DateFin := StrToDate(GetControlText('DATEFIN_'));
       StWhere := StWhere + ' AND AFF_DATEFIN >= "'+UsDateTime(DateDeb) +'" AND AFF_DATEFIN <= "'+UsDateTime(DateFin)+'"';
       end;

   for i:=1 to 3 do
      begin
      if (GetControlText('DATELIBRE'+intToStr(i)) <> DateTimeToStr(iDate1900))
         and (GetControlText('DATELIBRE'+intToStr(i))+'_' <> DateTimeToStr(iDate2099)) then
         begin
         DateDeb := StrToDate(GetControlText('DATELIBRE'+intToStr(i)));
         DateFin := StrToDate(GetControlText('DATELIBRE'+intToStr(i)+'_'));
         StWhere := StWhere + ' AND AFF_DATELIBRE'+intToStr(i)+' >= "'+UsDateTime(DateDeb) +'" AND AFF_DATELIBRE'+intToStr(i)+' <= "'+UsDateTime(DateFin)+'"';
         end;
      end;

   for i:=1 to 3 do
      begin
      if (GetCheckBoxState('BOOLLIBRE'+intToStr(i)) <> cbGrayed) then
         if GetCheckBoxState('BOOLLIBRE'+intToStr(i)) = cbChecked then
            StWhere:=StWhere + ' AND AFF_BOOLLIBRE'+intToStr(i)+' = "X"'
         else
            StWhere:=StWhere + ' AND AFF_BOOLLIBRE'+intToStr(i)+' = "-"';
      end;
   for i:=1 to 3 do
      begin
      if (Valeur(GetControlText('VALLIBRE'+intToStr(i))) <> 0)
         and (Valeur(GetControlText('VALLIBRE'+intToStr(i)+'_')) <> 0) then
         begin
         ValDeb := Valeur(GetControlText('VALLIBRE'+intToStr(i)));
         ValFin := Valeur(GetControlText('VALLIBRE'+intToStr(i))+'_');
         StWhere := StWhere + ' AND AFF_VALLIBRE'+intToStr(i)+' >= '+floatToStr(ValDeb) +' AND AFF_VALLIBRE'+intToStr(i)+' <= "'+floatToStr(ValFin)+'"';
         end;
      end;
   if StWhere<>'' then StWhere:=StWhere+'))';

   if (GetControlText('XX_WHERE') = '') then
      SetControlText('XX_WHERE',StWhere)
   else
       begin
       xx_where := GetControlText('XX_WHERE');
       xx_where := xx_where + ' and (' + StWhere + ')';
       SetControlText('XX_WHERE',xx_where) ;
       end;

    StWhere := '';
    if (GetControlText('COURRIER') <> 'X') then
       begin
         if GetControlText('FAX') = 'X' then StWhere := '((T_FAX <> "") or (C_FAX <> "")) and T_PARTICULIER <> "X"';
         if GetControlText('NOTFAX') = 'X' then StWhere := '(((T_FAX = "") or (T_FAX is null)) and ((C_FAX = "") or (C_FAX is null))) or T_PARTICULIER = "X"';
         if (GetControlText('XX_WHERE') = '') then
            SetControlText('XX_WHERE',StWhere)
         else
             begin
             xx_where := GetControlText('XX_WHERE');
             xx_where := xx_where + ' and (' + StWhere + ')';
             SetControlText('XX_WHERE',xx_where) ;
             end;
       end;

    if (GetControlText('XX_WHERE') = '') then
        SetControlText('XX_WHERE',RTXXWhereConfident('CON'))
    else
        begin
        xx_where := GetControlText('XX_WHERE');
        xx_where := xx_where + RTXXWhereConfident('CON');
        SetControlText('XX_WHERE',xx_where) ;
        end;
end;

Function TOF_MailingAff.MailingValidDoc : Boolean;
var stDocWord : string;
begin
Result:=False;
stDocWord:=getControlText('DOCGENERE');
if (stDocWord='') then begin PGIBox('Vous devez mentionner un document', ecran.caption); exit; end;
if not DirectoryExists(ExtractFilePath(stDocWord)) then begin PGIBox(Format(TraduireMemoire('Le répertoire %s n''existe pas'),[ExtractFilePath(stDocWord)]), ecran.caption); exit; end;
if FileExists(stDocWord) then
   if PGIAsk('Ce fichier existe déjà, voulez-vous l''écraser ?',ecran.caption)=mryes then DeleteFile(stDocWord ) else exit;
Result:=True;
end;

{$ifndef BUREAU}
procedure TOF_MailingAff.OpenDoc_OnClick(Sender: TObject) ;
var OkOk : Boolean;
begin
  AFLanceFiche_DocExterneMul('', 'TYPE=WOR;NATURE='+stNatureDoc+';FORM='+intToStr(longint(ecran)));
  if (GetControlText('MAQUETTE') <> '') and (stMaquetteValide <> '') and (not FileExists(stMaquetteValide)) then
    begin
    StMaquetteValide := '';
    OkOk := Ctrl_Maquette (GetControlText('MAQUETTE'),Ecran.Caption,stNatureDoc,StMaquetteValide);
    SetControlEnabled('BMAILINGDOC',  OkOk) ;
    SetControlEnabled('BEMAILING',  OkOk) ;
{$IF Defined(CRM)}
    SetControlEnabled('BESKERCOURRIER',OkOk and GetParamsocSecur('SO_RTGESTIONFLYDOC',False)) ;
    SetControlEnabled('BESKEREMAIL',OkOk and GetParamsocSecur('SO_RTGESTIONFLYDOC',False)) ;
    SetControlEnabled('BESKERFAX',OkOk and GetParamsocSecur('SO_RTGESTIONFLYDOC',False)) ;
{$IFEND}
    end;
end ;
{$endif}

{$ifndef BUREAU}
procedure TOF_MailingAff.MAQUETTE_OnElipsisClick(Sender: TObject);
var NoMaquette : string;
begin
  NoMaquette := AFLanceFiche_DocExterneMulSelec('ADE_DOCEXETAT=UTI', 'TYPE=WOR;NATURE='+stNatureDoc+';SELECTION');
  SetControlText('MAQUETTE',NoMaquette)
end;
{$endif}

{$ifndef BUREAU}
procedure TOF_MailingAff.ChangeMaquette(Sender: TObject);
var OkOk : Boolean;
begin
if StMaquetteValide <> '' then DeleteFile(pchar(StMaquetteValide));
StMaquetteValide := '';
OkOk := Ctrl_Maquette (GetControlText('MAQUETTE'),Ecran.Caption,stNatureDoc,StMaquetteValide);
SetControlEnabled('BMAILINGDOC',  OkOk) ;
SetControlEnabled('BEMAILING',  OkOk) ;
{$IF Defined(CRM)}
SetControlEnabled('BESKERCOURRIER',OkOk and GetParamsocSecur('SO_RTGESTIONFLYDOC',False)) ;
SetControlEnabled('BESKEREMAIL',OkOk and GetParamsocSecur('SO_RTGESTIONFLYDOC',False)) ;
SetControlEnabled('BESKERFAX',OkOk and GetParamsocSecur('SO_RTGESTIONFLYDOC',False)) ;
{$IFEND}
end;
{$endif}

{$ifndef BUREAU}
procedure TOF_MailingAff.BMAILINGDOC_OnClick(Sender: TObject) ;
var StTitre,stDocWord,stDocument : string;
    DS : THquery;
    OkOk : boolean ;
begin
if MailingValidDoc() = False then exit;
stDocWord:=getControlText('DOCGENERE');
{if (stDocWord='') then begin PGIBox('Vous devez mentionner un document', ecran.caption); exit; end;
if not DirectoryExists(ExtractFilePath(stDocWord)) then begin PGIBox(Format(TraduireMemoire('Le répertoire %s n''existe pas'),[ExtractFilePath(stDocWord)]), ecran.caption); exit; end;
if FileExists(stDocWord) then
   if PGIAsk('Ce fichier existe déjà, voulez-vous l''écraser ?',ecran.caption)=mryes then DeleteFile(stDocWord ) else exit;  }
{$IFDEF EAGLCLIENT}
if TFMul(Ecran).Fetchlestous then
   begin
{$ENDIF} //EAGLCLIENT
   DS:=THquery(TFMul(Ecran).Q) ;
   StTitre :=  TFMul(Ecran).Q.titres ;// GetTitreGrid (TFMul(Ecran).FListe);
   LancePublipostage('FILE',stMaquetteValide,stDocWord,Nil,StTitre,DS,True);
{$IFDEF EAGLCLIENT}
   end;
{$ENDIF}
OkOk:=False;
stDocument:=getControlText('DOCGENERE');
if (stDocument<>'') and (pos('\*.',stDocument)=0) and  FileExists(stDocument) then OkOk:=true;
SetControlEnabled('BDOCUMENT',OkOk) ;
end ;
{$endif}

procedure TOF_MailingAff.BEMAILING_OnClick(Sender: TObject);
var StTitre : string;
    DS : THquery;
begin
{$IFDEF EAGLCLIENT}
if Not TFMul(Ecran).Fetchlestous then exit;
{$ENDIF}
StTitre :=  TFMul(Ecran).Q.titres ;
DS := THquery(TFMul(Ecran).Q) ;
LancePublipostage('EMAIL',StMaquetteValide,'',Nil,StTitre,DS,False);
end;

procedure TOF_MailingAff.OnClose ;
begin
if StMaquetteValide <> '' then DeleteFile(pchar(StMaquetteValide));
end;

{$IFDEF AFFAIRE}
procedure TOF_MailingAff.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFFAIRE'));    Aff1:=THEdit(GetControl('AFFAIRE1')); Aff2:=THEdit(GetControl('AFFAIRE2'));
Aff3:=THEdit(GetControl('AFFAIRE3')); Aff4:=THEdit(GetControl('AVENANT'));
end;

procedure TOF_MailingAff.bSelectAff1Click(Sender: TObject);
begin
{$IFNDEF BTP}
    SelectionAffaire (EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4, VH_GC.GASeria , false, '', false, true, true)
{$ELSE}
    SelectionAffaire (EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4)
{$ENDIF}
end;
{$ENDIF AFFAIRE}

{$IF Defined(CRM)}
function TOF_MailingAff.ChpsObligExiste : Boolean;
begin
  Result := False;
  if TypeEnvoiEsker = 'MAIL' then
    begin
    if assigned(TFmul(Ecran).Q.FindField('C_RVA') ) and
      assigned(TFmul(Ecran).Q.FindField('T_EMAIL') ) then Result := True;
    end;
  if TypeEnvoiEsker = 'FAX' then
    begin
    if assigned(TFmul(Ecran).Q.FindField('C_FAX') ) and
      assigned(TFmul(Ecran).Q.FindField('T_FAX') ) then Result := True;
    end;
end;

procedure TOF_MailingAff.BESKERCOURRIER_OnClick(Sender: TObject) ;
var Infos : string;
    StTitre : string;
    DS : THquery;
begin
  if TFMul(Ecran).Q.RecordCount = 0 then exit;
  TypeEnvoiEsker := 'COURRIER';
  if MailingValidDoc() = False then Exit;
{$IFDEF EAGLCLIENT}
  if Not TFMul(Ecran).Fetchlestous then exit;
{$ENDIF}
  if LoginEsker = False then Exit;
  TobAdr:=Tob.create('LesAdresses',Nil , -1);
// Préparation de la Pré-Visualisation
  TFMul(Ecran).Q.First;
  TobAdr.AddChampSupValeur ('ADR_JURIDIQUE',TFmul(Ecran).Q.FindField('T_JURIDIQUE').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_LIBELLE',TFmul(Ecran).Q.FindField('T_LIBELLE').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_SUITELIBELLE',TFmul(Ecran).Q.FindField('T_PRENOM').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_ADRESSE1',TFmul(Ecran).Q.FindField('T_ADRESSE1').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_ADRESSE2',TFmul(Ecran).Q.FindField('T_ADRESSE2').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_ADRESSE3',TFmul(Ecran).Q.FindField('T_ADRESSE3').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_CODEPOSTAL',TFmul(Ecran).Q.FindField('T_CODEPOSTAL').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_VILLE',TFmul(Ecran).Q.FindField('T_VILLE').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_PAYS',TFmul(Ecran).Q.FindField('T_PAYS').asstring,False);
  TobAdr.AddChampSupValeur ('C_NOM',TFmul(Ecran).Q.FindField('C_NOM').asstring,False);
  TobAdr.AddChampSupValeur ('C_PRENOM',TFmul(Ecran).Q.FindField('C_PRENOM').asstring,False);
  TypeTraitement := 'PREVISU';
  StTitre :=  TFMul(Ecran).Q.titres ;
  DS := THquery(TFMul(Ecran).Q) ;
  LancePublipostage('SEPARATE',StMaquetteValide,GetControlText('DOCGENERE'),Nil,StTitre,DS,True,'',PublipostageCallBack);
  TobInfos:=Tob.create('LesInfos',Nil , -1);
  TobInfos.AddChampSupValeur ('TYPECOURRIER','M',False);
  Infos := RTLanceFiche_RTSaisieInfosEsker ('RT','RTSAISIECOURRIER','','','Document='+GetControlText('DOCGENERE')+';TobAdr='+intToStr(longint(TobAdr))+';TobInfos='+intToStr(longint(TobInfos)));
  if Infos <> '' then
  begin
    TobProsp:=Tob.create('_Les Prospects',Nil , -1);
    CreatTobInfosMul;
    InitMoveProgressForm(nil,Ecran.caption, 'Traitement en cours, veuillez patienter...', TFMul(Ecran).Q.RecordCount, True, True) ;
    TypeTraitement := 'REEL';
    LancePublipostage('SEPARATE',StMaquetteValide,GetControlText('DOCGENERE'),Nil,StTitre,DS,True,'',PublipostageCallBack);
    FiniMoveProgressForm ;
    FreeAndNil (TobProsp);
  end;
  LogoutEsker;
  FreeAndNil (TobInfos);
  FreeAndNil (TobAdr);
end ;

procedure TOF_MailingAff.BESKEREMAIL_OnClick(Sender: TObject) ;
var Infos,Chp : string;
    StTitre : string;
    DS : THquery;
    Q : TQuery;
begin
  if TFMul(Ecran).Q.RecordCount = 0 then exit;
  TypeEnvoiEsker := 'MAIL';
  if MailingValidDoc() = False then Exit;
  if ChpsObligExiste = False then
    begin
    PGIBox (TexteMessage[1],ecran.caption);
    Exit;
    end;
{$IFDEF EAGLCLIENT}
  if Not TFMul(Ecran).Fetchlestous then exit;
{$ENDIF}
  if LoginEsker = False then Exit;
  TobAdr:=Tob.create('LesAdresses',Nil , -1);
  TobInfos:=Tob.create('LesInfos',Nil , -1);
  TobInfos.AddChampSupValeur ('TYPECOURRIER','M',False);
  Infos := RTLanceFiche_RTSaisieInfosEsker ('RT','RTSAISIEEMAIL','','','TobInfos='+intToStr(longint(TobInfos)));
  if Infos <> '' then
  begin
    Q:=OpenSQL('Select ARS_EMAIL FROM RESSOURCE WHERE ARS_UTILASSOCIE="'+V_PGI.User+'"',true);
    if Not Q.EOF then Chp:=Q.FindField('ARS_EMAIL').asstring;
    ferme(Q);
    TobInfos.AddChampSupValeur ('EMAILEMETTEUR',Chp,False);
    TobAdr.AddChampSup ('NOMCONTACT',False);
    TobAdr.AddChampSup ('NOMSOCCONTACT',False);
    TobAdr.AddChampSup ('EMAILCONTACT',False);
    TobProsp:=Tob.create('_Les Prospects',Nil , -1);
    CreatTobInfosMul;
    StTitre :=  TFMul(Ecran).Q.titres ;
    DS := THquery(TFMul(Ecran).Q) ;
    InitMoveProgressForm(nil,Ecran.caption, 'Traitement en cours, veuillez patienter...', TFMul(Ecran).Q.RecordCount, True, True) ;
    LancePublipostage('SEPARATE',StMaquetteValide,GetControlText('DOCGENERE'),Nil,StTitre,DS,True,'',PublipostageCallBack);
    FiniMoveProgressForm ;
    FreeAndNil (TobProsp);
  end;
  LogoutEsker;
  FreeAndNil (TobInfos);
  FreeAndNil (TobAdr);
end ;

procedure TOF_MailingAff.BESKERFAX_OnClick(Sender: TObject) ;
var Infos : string;
    StTitre : string;
    DS : THquery;
    Chp : string;
begin
  if TFMul(Ecran).Q.RecordCount = 0 then exit;
  TypeEnvoiEsker := 'FAX';
  if MailingValidDoc() = False then Exit;
  if ChpsObligExiste = False then
    begin
    PGIBox (TexteMessage[2],ecran.caption);
    Exit;
    end;
{$IFDEF EAGLCLIENT}
  if Not TFMul(Ecran).Fetchlestous then exit;
{$ENDIF}
  if LoginEsker = False then Exit;
// Préparation de la Pré-Visualisation
  TypeTraitement := 'PREVISU';
  StTitre :=  TFMul(Ecran).Q.titres ;
  DS := THquery(TFMul(Ecran).Q) ;
  LancePublipostage('SEPARATE',StMaquetteValide,GetControlText('DOCGENERE'),Nil,StTitre,DS,True,'',PublipostageCallBack);
  TobInfos:=Tob.create('LesInfos',Nil , -1);
  TobInfos.AddChampSupValeur ('TYPECOURRIER','M',False);
  TobAdr:=Tob.create('LesAdresses',Nil , -1);
  TFMul(Ecran).Q.First;
  TobAdr.AddChampSup ('NOMCONTACT',False);
  TobAdr.AddChampSup ('NOMSOCCONTACT',False);
  TobAdr.AddChampSup ('FAXCONTACT',False);
  TobAdr.SetString ('NOMCONTACT',TFmul(Ecran).Q.FindField('C_PRENOM').asstring + ' ' + TFmul(Ecran).Q.FindField('C_NOM').asstring);
  TobAdr.SetString ('NOMSOCCONTACT',TFmul(Ecran).Q.FindField('T_LIBELLE').asstring);
  if TFmul(Ecran).Q.FindField('C_FAX').asstring <> '' then Chp := TFmul(Ecran).Q.FindField('C_FAX').asstring
  else Chp := TFmul(Ecran).Q.FindField('T_FAX').asstring;
  TobAdr.SetString ('FAXCONTACT',CleTelephone (Chp));
  Infos := RTLanceFiche_RTSaisieInfosEsker ('RT','RTSAISIEFAX','','','Document='+GetControlText('DOCGENERE')+';TobAdr='+intToStr(longint(TobAdr))+';TobInfos='+intToStr(longint(TobInfos)));
  if Infos <> '' then
  begin
    TobProsp:=Tob.create('_Les Prospects',Nil , -1);
    CreatTobInfosMul;
    InitMoveProgressForm(nil,Ecran.caption, 'Traitement en cours, veuillez patienter...', TFMul(Ecran).Q.RecordCount, True, True) ;
    TypeTraitement := 'REEL';
    LancePublipostage('SEPARATE',StMaquetteValide,GetControlText('DOCGENERE'),Nil,StTitre,DS,True,'',PublipostageCallBack);
    FiniMoveProgressForm ;
    FreeAndNil (TobProsp);
  end;
  LogoutEsker;
  FreeAndNil (TobInfos);
  FreeAndNil (TobAdr);
end ;

procedure TOF_MailingAff.PublipostageCallBack(Index: integer; OutputFileName: hstring; var Cancel: boolean);
begin
if TypeTraitement = 'PREVISU' then
  begin
  Cancel := True;
  exit;
  end;
if Not MoveCurProgressForm(format(TexteMessage[3],[IntToStr(Index+1)])) then
  begin
  Cancel := True;
  exit;
  end
else
  begin
  Cancel := false;
  if FileExists(OutputFileName) then
    begin
    if TypeEnvoiEsker = 'COURRIER' then
      begin
      TobAdr.SetString ('ADR_JURIDIQUE',TobProsp.detail[index].GetString('ADR_JURIDIQUE'));
      TobAdr.SetString ('ADR_LIBELLE',TobProsp.detail[index].GetString('ADR_LIBELLE'));
      TobAdr.SetString ('ADR_SUITELIBELLE',TobProsp.detail[index].GetString('ADR_SUITELIBELLE'));
      TobAdr.SetString ('ADR_ADRESSE1',TobProsp.detail[index].GetString('ADR_ADRESSE1'));
      TobAdr.SetString ('ADR_ADRESSE2',TobProsp.detail[index].GetString('ADR_ADRESSE2'));
      TobAdr.SetString ('ADR_ADRESSE3',TobProsp.detail[index].GetString('ADR_ADRESSE3'));
      TobAdr.SetString ('ADR_CODEPOSTAL',TobProsp.detail[index].GetString('ADR_CODEPOSTAL'));
      TobAdr.SetString ('ADR_VILLE',TobProsp.detail[index].GetString('ADR_VILLE'));
      TobAdr.SetString ('ADR_PAYS',TobProsp.detail[index].GetString('ADR_PAYS'));
      TobAdr.SetString ('C_NOM',TobProsp.detail[index].GetString('C_NOM'));
      TobAdr.SetString ('C_PRENOM',TobProsp.detail[index].GetString('C_PRENOM'));
      SendEsker (0,TobInfos,OutputFileName,TobAdr);
      end else
      if TypeEnvoiEsker = 'MAIL' then
        begin
        if TobProsp.detail[index].GetString('EMAIL') <> '' then
          begin
          TobAdr.SetString ('NOMCONTACT',TobProsp.detail[index].GetString('C_NOM'));
          TobAdr.SetString ('NOMSOCCONTACT',TobProsp.detail[index].GetString('ADR_LIBELLE'));
          TobAdr.SetString ('EMAILCONTACT',TobProsp.detail[index].GetString('EMAIL'));
          SendEsker (2,TobInfos,OutputFileName,TobAdr);
          end;
        end else
        if TypeEnvoiEsker = 'FAX' then
          begin
          if TobProsp.detail[index].GetString('FAX') <> '' then
            begin
            TobAdr.SetString ('NOMCONTACT',TobProsp.detail[index].GetString('C_PRENOM') + ' ' + TobProsp.detail[index].GetString('C_NOM'));
            TobAdr.SetString ('NOMSOCCONTACT',TobProsp.detail[index].GetString('ADR_LIBELLE'));
            TobAdr.SetString ('FAXCONTACT',TobProsp.detail[index].GetString('FAX'));
            SendEsker (1,TobInfos,OutputFileName,TobAdr);
            end;
          end;
    DeleteFile(OutputFileName )
    end;
  end;
end;

procedure TOF_MailingAff.CreatTobInfosMul ;
var Q : THQuery;
    TobAdrCli : Tob;
    Chp : string;
begin
  Q:= TFMul(Ecran).Q;
  InitMove(Q.RecordCount,'');
  Q.First;
  while Not Q.EOF do
    begin
    MoveCur(False);
    TobAdrCli := Tob.create('',TobProsp,-1);
    TobAdrCli.AddChampSupValeur ('ADR_JURIDIQUE',TFmul(Ecran).Q.FindField('T_JURIDIQUE').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_LIBELLE',TFmul(Ecran).Q.FindField('T_LIBELLE').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_SUITELIBELLE',TFmul(Ecran).Q.FindField('T_PRENOM').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_ADRESSE1',TFmul(Ecran).Q.FindField('T_ADRESSE1').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_ADRESSE2',TFmul(Ecran).Q.FindField('T_ADRESSE2').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_ADRESSE3',TFmul(Ecran).Q.FindField('T_ADRESSE3').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_CODEPOSTAL',TFmul(Ecran).Q.FindField('T_CODEPOSTAL').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_VILLE',TFmul(Ecran).Q.FindField('T_VILLE').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_PAYS',TFmul(Ecran).Q.FindField('T_PAYS').asstring,False);
    TobAdrCli.AddChampSupValeur ('C_NOM',TFmul(Ecran).Q.FindField('C_NOM').asstring,False);
    TobAdrCli.AddChampSupValeur ('C_PRENOM',TFmul(Ecran).Q.FindField('C_PRENOM').asstring,False);
    if TypeEnvoiEsker = 'MAIL' then
      begin
      if TFmul(Ecran).Q.FindField('C_RVA').asstring <> '' then Chp := TFmul(Ecran).Q.FindField('C_RVA').asstring
      else Chp := TFmul(Ecran).Q.FindField('T_EMAIL').asstring;
      TobAdrCli.AddChampSupValeur ('EMAIL',Chp,False);
      end;
    if TypeEnvoiEsker = 'FAX' then
      begin
      if TFmul(Ecran).Q.FindField('C_FAX').asstring <> '' then Chp := TFmul(Ecran).Q.FindField('C_FAX').asstring
      else Chp := TFmul(Ecran).Q.FindField('T_FAX').asstring;
      TobAdrCli.AddChampSupValeur ('FAX',CleTelephone (Chp),False);
      end;
    Q.Next;
    end;
end;

procedure TOF_MailingAff.BESKERSUIVI_OnClick(Sender: TObject) ;
var sHttp : String;
begin
  sHttp := 'https://as1.ondemand.esker.com/ondemand/webaccess/40/documents.aspx';
  LanceWeb(sHttp,False);
end;

{$IFEND}


procedure TOF_MailingAff.constitueXXWHEREPARC;
var TheXXWhere : THEdit;
    TheWhere : string;
begin
  TheXXWhere := THEdit(GetControl('XX_WHERE1'));
  if TheXXWhere = nil then Exit;
  TheXXWhere.text := '';
  if (GetControlText('TYPEARTICLE')='') and
     (GetControlText('CODEARTICLE')='') and
     (GetControlText('REFFABRICANT')='') and
     (GetControlText('SERIE')='') and
     (GetControlText('VERSION')='') and
     (GetControlText('DATEFINSERIA')='31/12/2099') and
     (GetControlText('REFCLIENT')='') then Exit;

  TheWhere := ' AND EXISTS (SELECT BP1_ARTICLE FROM PARCTIERS WHERE BP1_TIERS=T_TIERS ';
  if GetControlText('TYPEARTICLE')<>'' then TheWhere := TheWhere+ 'AND BP1_TYPEARTICLE="'+GetControlText('TYPEARTICLE')+'" ';
  if GetControlText('CODEARTICLE')<>'' then TheWhere := TheWhere+ 'AND BP1_CODEARTICLE LIKE "'+GetControlText('CODEARTICLE')+'%" ';
  if GetControlText('REFFABRICANT')<>'' then TheWhere := TheWhere+ 'AND BP1_REFFABRICANT LIKE "'+GetControlText('REFFABRICANT')+'%" ';
  if GetControlText('REFCLIENT')<>'' then TheWhere := TheWhere+ 'AND BP1_REFCLIENT LIKE "'+GetControlText('REFCLIENT')+'%" ';
  if GetControlText('SERIE')<>'' then TheWhere := TheWhere+ 'AND BP1_SERIE LIKE "'+GetControlText('SERIE')+'%" ';
  if GetControlText('VERSION')<>'' then TheWhere := TheWhere+ 'AND BP1_CODEVERSION LIKE "'+GetControlText('VERSION')+'%" ';
  if GetControlText('DATEFINSERIA')<>'31/12/2099' then TheWhere := TheWhere+ 'AND BP1_DATEFINSERIA < "'+USDATETIME(StrToDate(GetControlText('DATEFINSERIA')))+'"';
  TheXXWhere.text := TheWhere + ' AND BP1_ETATPARC="ES")';

end;

procedure TOF_MailingAff.BVALIDSELClick(Sender: TObject);
begin
  constitueXXWHEREPARC;
  BchercheClick(self);
end;

procedure TOF_MailingAff.CodeArticleRech(Sendre: TOBject);
var stWhere,stFiche : string;
		ART : THCritMaskEdit;
    CODEARTICLE : Thedit;
begin
  CODEARTICLE := ThEdit(GetControl('CODEARTICLE'));
	ART := THCritMaskEdit.Create (ecran); ART.Visible := false;
  ART.Text := '';
  //
  stFiche := 'BTARTPARC_RECH';
  if GetControlText('TYPEARTICLE') <> '' then
  begin
	  stWhere := 'GA_TYPEARTICLE = "'+GetControlText('TYPEARTICLE')+'"';
  end else
  begin
	  stWhere := 'GA_TYPEARTICLE IN ("PA1","PA2")';
  end;

  StWhere := 'GA_CODEARTICLE=' + Trim (Copy (ART.Text, 1, 18))+';XX_WHERE=AND '+stWhere;
  DispatchRecherche (ART, 1, '',stWhere, stFiche);
	if ART.Text <> '' then
  	CODEARTICLE.Text := Trim(copy(ART.Text, 0, Length(ART.Text) - 1));
  ART.free;
end;


procedure TOF_MailingAff.ScreenKeYDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 120 then
  begin
    TToolbarButton97(GetControl('BVALIDSEL')).Click;
    Key := 0;
  end;
end;


Initialization
registerclasses([TOF_MailingAff]);
end.
