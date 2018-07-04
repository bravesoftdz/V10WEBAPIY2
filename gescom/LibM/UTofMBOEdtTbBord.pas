unit UTofMBOEdtTbBord;

interface

uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, UTOB,Spin,UtilSynVte,
{$IFDEF EAGLCLIENT}
      eQRS1,utileAGL,
{$ELSE}
      QRS1,db,dbTables,EdtEtat,Fiche,
{$ENDIF}
      AglInit,EntGC,HQry,HStatus,MajTable,UtilGc,graphics,M3FP,VoirTob ;

Type
     TOF_MBOEdtTbBord = Class (TOF)
        procedure OnUpdate ; override ;
        procedure OnClose ; override ;
        procedure OnArgument (Argument : String ) ; override ;
     private
        SemaineFin,NumSemaineFinAnneePrec,MoisFin : integer ; // Date de fin exprimée en semaine, mois
        procedure RempliTOBTbBordStock(var TobTbBord : TOB; TobStock : TOB; CodeTrait : string) ;
        procedure RempliTOBTbBordVente(var TobTbBord : TOB; TobVentes : TOB; CodeTrait : string) ;
        function initChampsTOB(T : TOB) : integer ;
     end;


implementation

procedure TOF_MBOEdtTbBord.OnArgument (Argument : String ) ;
var iInd : integer ;
begin
inherited ;

for iInd:=1 to 3 do ChangeLibre2('TGA_FAMILLENIV'+IntToStr(iInd),Ecran);
// Récupération des libellés des 3 champs libre article -> impression titre des colonnes
GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GA_LIBREART', 3, '');
// paramétrage du libellé etablissement quand on est en multi-dépôts
if VH_GC.GCMultiDepots then
   begin
   SetControlText('TDEPOT',TraduireMemoire('Dépôts'));
   SetControlCaption('DETAILETAB',TraduireMemoire('Détaillé par dépôt'));
   THValComboBox (TForm(Ecran).FindComponent('RUPT1')).Plus := ' AND CO_CODE<>"ETA" ';
   THValComboBox (TForm(Ecran).FindComponent('RUPT2')).Plus := ' AND CO_CODE<>"ETA" ';
   THValComboBox (TForm(Ecran).FindComponent('RUPT3')).Plus := ' AND CO_CODE<>"ETA" ';
   end
else
   begin
   SetControlText('TDEPOT',TraduireMemoire('Etablissement'));
   SetControlCaption('DETAILETAB',TraduireMemoire('Détaillé par établissement'));
   THValComboBox (TForm(Ecran).FindComponent('RUPT1')).Plus := ' AND CO_CODE<>"DEP" ';
   THValComboBox (TForm(Ecran).FindComponent('RUPT2')).Plus := ' AND CO_CODE<>"DEP" ';
   THValComboBox (TForm(Ecran).FindComponent('RUPT3')).Plus := ' AND CO_CODE<>"DEP" ';
   end;
end;

procedure TOF_MBOEdtTbBord.OnUpdate ;
var bFirstCrit : boolean ;
    iInd,i : integer ;
    stWhereStock,stWherePiece,stCrit,stTmp,Sql,SqlGroup : string ;
    QQ : TQuery ;
    TobTbBord,TobStock,TobVentes : Tob ;
begin
Inherited;

ExecuteSQL('delete from GCTMPTABBOR where GZV_UTILISATEUR = "'+V_PGI.USer+'" and GZV_TRAIT = "CDE"') ;


// 1. Chargement des articles et qtés en stock
// ===========================================
stWhereStock:=' where GQ_CLOTURE<>"X"' ;

stCrit:=GetControlText('CODEARTICLE') ;
if stCrit<>'' then stWhereStock:=stWhereStock+' and GA_CODEARTICLE>="'+stCrit+'"' ;
stCrit:=GetControlText('CODEARTICLE_') ;
if stCrit<>'' then stWhereStock:=stWhereStock+' and GA_CODEARTICLE<="'+stCrit+'"' ;
stCrit:=GetControlText('FOURNPRINC') ;
if stCrit<>'' then stWhereStock:=stWhereStock+' and GA_FOURNPRINC>="'+stCrit+'"' ;
stCrit:=GetControlText('FOURNPRINC_') ;
if stCrit<>'' then stWhereStock:=stWhereStock+' and GA_FOURNPRINC<="'+stCrit+'"' ;
for iInd:=1 to 3 do
    BEGIN
    stCrit:=GetControlText('GA_FAMILLENIV'+IntToStr(iInd)) ;
    if stCrit<>'' then stWhereStock:=stWhereStock+' and GA_FAMILLENIV'+IntToStr(iInd)+'="'+stCrit+'"' ;
    END ;

stCrit := GetControlText('DEPOT') ;
if (stCrit <> '') and (stCrit <> TraduireMemoire ('<<Tous>>')) then
    BEGIN
    bFirstCrit:=True ;
    stWhereStock:=stWhereStock+' and GQ_DEPOT in (' ;
    repeat
        stTmp:=ReadTokenSt(stCrit) ;
        if bFirstCrit then bFirstCrit:=False else stWhereStock:=stWhereStock+',' ;
        stWhereStock:=stWhereStock+'"'+stTmp+'"' ;
    until stCrit='' ;
    stWhereStock:=stWhereStock+')' ;
    END ;
if GetControlText('CRUPTURE')='X' then stWhereStock:=stWhereStock+' and (GQ_PHYSIQUE-GQ_RESERVECLI+GQ_RESERVEFOU)<GQ_STOCKMIN'
else if GetControlText('CSURSTOCK')='X' then stWhereStock:=stWhereStock+' and (GQ_PHYSIQUE-GQ_RESERVECLI+GQ_RESERVEFOU)>GQ_STOCKMAX' ;


// Construction de la requête
// ==========================
// (stWhereStock contient au moins la clause GQ_CLOTURE<>"X")
Sql:='select GA_CODEARTICLE as GZV_CODEARTICLE,' ;
if GetControlText('DETAILETAB')='X' then Sql:=Sql+'GQ_DEPOT as GZV_DEPOT, ' ;
Sql:=Sql+'GA_FOURNPRINC as GZV_FOURNPRINC,GA_FAMILLENIV1 as GZV_FAMILLENIV1,' +
         'GA_FAMILLENIV2 as GZV_FAMILLENIV2,GA_FAMILLENIV3 as GZV_FAMILLENIV3,' ;
SqlGroup:='' ;
for iInd:=1 to 3 do
    BEGIN
    StTmp:=IntToStr(iInd) ;
    stCrit:=GetControlText('RUPTURE'+StTmp) ;
    if stCrit<>'' then
        BEGIN
        if stCrit='GQ_DEPOT' then Sql:=Sql+'GQ_DEPOT as GZV_DEPOT, ' ;
        Sql:=Sql+StCrit+' as GZV_RUPT'+StTmp+',' ;
        SqlGroup:=SqlGroup+','+StCrit ;
        END ;
    END;

Sql:=Sql+'sum(GQ_PHYSIQUE) as GZV_PHYSIQUE,sum(GQ_RESERVEFOU) as GZV_RESERVEFOU' +
         ' from DISPO left join ARTICLE on GQ_ARTICLE=GA_ARTICLE ' +
         stWhereStock +
         ' group by GA_CODEARTICLE,' ;
if GetControlText('DETAILETAB')='X' then Sql:=Sql+'GQ_DEPOT,' ;
Sql:=Sql+'GA_FOURNPRINC,GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3'+SqlGroup ;


QQ:=OpenSQL(Sql,True) ;
if QQ.Eof then
   begin
   Ferme(QQ);
   exit;
   end;
TobStock:=TOB.Create('LE STOCK',nil,-1) ;
TobStock.LoadDetailDB('GCTMPTABBOR','','',QQ,False,True) ;
Ferme(QQ) ;

TobTbBord:=TOB.Create('Tableau Bord',nil,-1) ;
RempliTOBTBBORDStock(TobTbBord,TobStock,'CDE') ;
TobStock.free;
// Initialise les champs GZV_UTILISATEUR, GZV_TRAIT, GZV_QTE0..12
TobTbBord.ParcoursTraitement([''],[''],False,initChampsTOB) ;


// 2. Chargement des qtés vendues par période
// ==========================================

stCrit := GetControlText('NATUREPIECEG') ;
if (stCrit <> '') and (stCrit <> TraduireMemoire ('<<Tous>>')) then
    BEGIN
    stWherePiece:=' where GL_NATUREPIECEG in (' ;
    bFirstCrit:=False ;
    repeat
        stTmp:=ReadTokenSt(stCrit) ;
        if bFirstCrit then stWherePiece:=stWherePiece+',' else bFirstCrit:=True ;
        stWherePiece:=stWherePiece+'"'+stTmp+'"' ;
    until stCrit='' ;
    stWherePiece:=stWherePiece+')' ;
    END
    else stWherePiece:=' where GL_NATUREPIECEG in ("AVC","AVS","FAC","FFO")' ;

stCrit:=GetControlText('DATEPIECEDEB') ;
if stCrit<>'' then stWherePiece:=stWherePiece+' and GL_DATEPIECE>="'+USDateTime(StrToDate(stCrit))+'"' ;
stCrit:=GetControlText('DATEPIECEFIN') ;
if stCrit<>'' then stWherePiece:=stWherePiece+' and GL_DATEPIECE<="'+USDateTime(StrToDate(stCrit))+'"' ;

stCrit:=GetControlText('CODEARTICLE') ;
if stCrit<>'' then stWherePiece:=stWherePiece+' and GL_CODEARTICLE>="'+stCrit+'"'
              else stWherePiece:=stWherePiece+' and GL_CODEARTICLE<>""' ;
stCrit:=GetControlText('CODEARTICLE_') ;
if stCrit<>'' then stWherePiece:=stWherePiece+' and GL_CODEARTICLE<="'+stCrit+'"' ;
for iInd:=1 to 3 do
    BEGIN
    stCrit:=GetControlText('GA_FAMILLENIV'+IntToStr(iInd)) ;
    if stCrit<>'' then stWherePiece:=stWherePiece+' and GL_FAMILLENIV'+IntToStr(iInd)+'="'+stCrit+'"' ;
    END ;

if GetControlText('DETAILETAB')='X' then
    BEGIN
    Sql:='select GL_CODEARTICLE,GL_DEPOT as GL_DEPOT, ' +
         DB_Week('GL_DATEPIECE')+' as DATEPIECE, sum(GL_QTEFACT) as QTEVENTE' +
         ' from LIGNE'+stWherePiece +
         ' group by GL_CODEARTICLE,GL_DEPOT,'+DB_Week('GL_DATEPIECE') ;
    END else
    BEGIN
    // j'ai mis "-" pour l'établissement pour que la colonne GL_ETABLISSEMENT soit prise en compte en ORACLE
    Sql:='select GL_CODEARTICLE,"-" as GL_DEPOT, ' +
         DB_Week('GL_DATEPIECE')+' as DATEPIECE, sum(GL_QTEFACT) as QTEVENTE' +
         ' from LIGNE'+stWherePiece +
         ' group by GL_CODEARTICLE,GL_DEPOT,'+DB_Week('GL_DATEPIECE') ;
    END ;

QQ:=OpenSQL(Sql,True) ;
if not QQ.Eof then
    BEGIN
    TobVentes:=TOB.Create('Les ventes',nil,-1) ;
    TobVentes.LoadDetailDB('','','',QQ,False,True) ;
    Ferme(QQ) ;
    // je remet "" dans GL_ETABLISSEMENT sur toutes les lignes
    if GetControlText('DETAILETAB')='-' then
    BEGIN
    for i:=0 to TobVentes.Detail.count-1 do
       begin
       TobVentes.Detail[i].PutValue('GL_DEPOT','');
       end;
    end;
    //case en fonction type période sélectionné
    SemaineFin:=NumSemaine(StrToDate(GetControlText('DATEPIECEFIN'))) ;
    // Numéro semaine du 24/12 : évite éventuelle semaine no 1 de l'année suivante
    NumSemaineFinAnneePrec:=NumSemaine(StrToDate('24/12/'+Copy(GetControlText('DATEPIECEDEB'),7,4))) ;
    MoisFin:=StrToInt(copy(GetControlText('DATEPIECEFIN'),4,2)) ;

    RempliTOBTBBORDVente(TobTbBord,TobVentes,'CDE') ;
    TobVentes.Free ;
    END ;

if TobTbBord.Detail.Count>0 then
    BEGIN
    TobTbBord.SetAllModifie(True) ;
    TobTbBord.InsertDB(Nil) ;
    END ;
TobTbBord.Free ;

SetControlText('XX_WHERE','GZV_UTILISATEUR="'+V_PGI.USer+'" and GZV_TRAIT="CDE"') ;

end;

procedure TOF_MBOEdtTbBord.OnClose ;
begin
Inherited;
ExecuteSQL('delete from GCTMPTABBOR where GZV_UTILISATEUR = "'+V_PGI.USer+'" and GZV_TRAIT = "CDE"');
end;

procedure TOF_MBOEdtTbBord.RempliTOBTbBordStock(var TobTbBord : TOB; TobStock : TOB; CodeTrait : string) ;
var TobPal,TobT : TOB ;
    i_ind,iRupt : integer ;
    stRupt : string ;
begin
InitMove(TobStock.Detail.count,'') ;
for i_ind:=0 to TobStock.Detail.count-1 do
    BEGIN
    TobT:=TobStock.Detail[i_ind] ;
    TobPal:=TOB.Create('GCTMPTABBOR',TobTbBord,-1) ;
    for iRupt:=1 to 3 do
        BEGIN
        stRupt:=IntToStr(iRupt) ;
        if TobT.FieldExists('GZV_RUPT'+stRupt) then TobPal.PutValue('GZV_RUPT'+stRupt,TobT.GetValue('GZV_RUPT'+stRupt))
        else TobPal.PutValue('GZV_RUPT'+stRupt,stRupt) ;
        END ;
    TobPal.PutValue('GZV_FOURNPRINC',TobT.GetValue('GZV_FOURNPRINC')) ;
    TobPal.PutValue('GZV_CODEARTICLE',TobT.GetValue('GZV_CODEARTICLE')) ;
    TobPal.PutValue('GZV_DEPOT',TobT.GetValue('GZV_DEPOT')) ;
    TobPal.PutValue('GZV_FAMILLENIV1',TobT.GetValue('GZV_FAMILLENIV1')) ;
    TobPal.PutValue('GZV_FAMILLENIV2',TobT.GetValue('GZV_FAMILLENIV2')) ;
    TobPal.PutValue('GZV_FAMILLENIV3',TobT.GetValue('GZV_FAMILLENIV3')) ;
    TobPal.PutValue('GZV_PHYSIQUE',TobT.GetValue('GZV_PHYSIQUE')) ;
    TobPal.PutValue('GZV_RESERVEFOU',TobT.GetValue('GZV_RESERVEFOU')) ;
    MoveCur(False) ;
    END ;
FiniMove ;
end ;

procedure TOF_MBOEdtTbBord.RempliTOBTbBordVente(var TobTbBord : TOB; TobVentes : TOB; CodeTrait : string) ;
var TobPal,TobT : TOB ;
    i_ind,iPer : integer ;
    stPer : string ;
    qteVte : double ;
begin
InitMove(TobVentes.Detail.count,'') ;
for i_ind:=0 to TobVentes.Detail.count-1 do
    BEGIN
    TobT:=TobVentes.Detail[i_ind] ;
    iPer:=SemaineFin-TobT.GetValue('DATEPIECE') ;
    if iPer<0 then iPer:=iPer+NumSemaineFinAnneePrec ; // Gestion changement d'année
    if (iPer>-1) and (iPer<13) then
        BEGIN
        TobPal:=TobTbBord.FindFirst(['GZV_UTILISATEUR','GZV_TRAIT','GZV_DEPOT','GZV_CODEARTICLE'],
                [V_PGI.User,CodeTrait,TobT.GetValue('GL_DEPOT'),TobT.GetValue('GL_CODEARTICLE')], False) ;
        if TobPal<>nil then
            BEGIN
            stPer:=IntToStr(iPer) ;
            if TobPal.GetValue('GZV_QTE'+stPer)<>null then qteVte:=TobPal.GetValue('GZV_QTE'+stPer) else qteVte:=0.0 ;
            if TobT.GetValue('QTEVENTE')<>null then qteVte:=qteVte+TobT.GetValue('QTEVENTE') ;
            TobPal.PutValue('GZV_QTE'+stPer,qteVte) ;
            TobPal.PutValue('GZV_CODEARTICLE',TobPal.GetValue('GZV_CODEARTICLE')) ;
            END ;
        END ;
    MoveCur(False) ;
    END ;
FiniMove ;
end ;

function TOF_MBOEdtTbBord.initChampsTOB(T : TOB) : integer ;
var iInd : integer ;
begin
// Initialise les champs GZV_UTILISATEUR, GZV_TRAIT, GZV_QTE0..12
T.PutValue('GZV_UTILISATEUR',V_PGI.User) ;
T.PutValue('GZV_TRAIT','CDE') ;
if T.GetValue('GZV_DEPOT')=null then T.PutValue('GZV_DEPOT','...') ;
for iInd:=0 to 12 do T.PutValue('GZV_QTE'+IntToStr(iInd),0.0) ;
T.PutValue('GZV_PROPOSCDE',0.0) ;
Result:=0 ;
end ;

/////////CONTROLE DES RUPTURES PARAMETRABLES////////////////////////////////////////
Procedure TOF_MBOEdtTbBord_AffectGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value, St_Text, St, StIndice : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
Indice := Integer (Parms[1]);
StIndice := IntToStr(Indice);
if VH_GC.GCMultiDepots then St_Plus := ' AND CO_CODE<>"ETA" '
else St_Plus := ' AND CO_CODE<>"DEP" ';
St_Value := string (THValComboBox (F.FindComponent('RUPT'+StIndice)).Value);
St_Text := string (THValComboBox (F.FindComponent('RUPT'+StIndice)).Text);
For i_ind := 1 to 3 do
    BEGIN
    if i_ind = Indice then continue;
    St := string (THValComboBox (F.FindComponent('RUPT'+IntToStr(i_ind))).Value);
    If St <> '' then St_Plus := St_Plus + ' AND CO_CODE <>"' + St + '"' ;
    END;
THValComboBox (F.FindComponent('RUPT'+StIndice)).Plus := St_Plus;
THValComboBox (F.FindComponent('RUPT'+StIndice)).Value := St_Value;
THValComboBox (F.FindComponent('RUPT'+StIndice)).Text := St_Text;
END;

Procedure TOF_MBOEdtTbBord_ChangeGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value, St_Indice, St_Ind : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
Indice := Integer (Parms[1]);
St_Indice := IntToStr(Indice) ;
if VH_GC.GCMultiDepots then St_Plus := ' AND CO_CODE<>"ETA" '
else St_Plus := ' AND CO_CODE<>"DEP" ';
St_Value := string (THValComboBox (F.FindComponent('RUPT'+St_Indice)).Value);
if St_Value = '' then
    BEGIN
    THEdit (F.FindComponent('RUPTURE'+St_Indice)).Text := '';
    THEdit (F.FindComponent('XX_VARIABLE'+St_Indice)).Text := '';
    THValComboBox (F.FindComponent('RUPT'+St_Indice)).Plus := St_Plus;
    THValComboBox (F.FindComponent('RUPT'+St_Indice)).Text := '';
    For i_ind := Indice + 1 to 3 do
        BEGIN
        St_Ind := IntToStr(i_ind) ;
        THValComboBox (F.FindComponent('RUPT'+St_Ind)).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+St_Ind)).Value := '';
        THValComboBox (F.FindComponent('RUPT'+St_Ind)).Plus := St_Plus;
        THValComboBox (F.FindComponent('RUPT'+St_Ind)).Enabled := False;
        THValComboBox (F.FindComponent('RUPT'+St_Ind)).Color := clBtnFace;
        THEdit (F.FindComponent('RUPTURE'+St_Ind)).Text := '';
        THEdit (F.FindComponent('XX_VARIABLE'+St_Ind)).Text := '';
        END;
    END else
    BEGIN
    if Indice < 3 then
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+IntToStr(Indice + 1))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+IntToStr(Indice + 1))).Color := clWindow;
        END;
    THEdit (F.FindComponent('RUPTURE'+St_Indice)).Text :=
            RechDom ('GCGROUPTABBOR', St_Value, True);
    THEdit (F.FindComponent('XX_VARIABLE'+St_Indice)).Text :=
            string (THValComboBox (F.FindComponent('RUPT'+St_Indice)).Text);
    END;
END;

procedure InitTOFTbBord ();
begin
RegisterAglProc('TbBord_ChangeGroup', True , 1, TOF_MBOEdtTbBord_ChangeGroup);
RegisterAglProc('TbBord_AffectGroup', True , 1, TOF_MBOEdtTbBord_AffectGroup);
end;

Initialization
registerclasses([TOF_MBOEdtTbBord]) ;
InitTOFTbBord();

end.
