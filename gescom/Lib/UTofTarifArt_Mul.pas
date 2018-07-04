Unit UTofTarifArt_Mul ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     emul,
{$ELSE}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBGrids,mul,
{$ENDIF}
     forms, sysutils,ComCtrls,
     HCtrls, HEnt1, HMsgBox, HTB97, UTOF,
     Hqry, UTOB,
//uniquement en line
//   BTASSIST_MAJ_TARIF,
     AssistMajTarifArt,
     HDimension, M3FP, UtilGC, AssistCreationTarifArt;

Type
  Tof_TarifArt_Mul = Class (TOF)
   public
    { Déclarations publiques }
    procedure OnArgument (Arguments : String ) ; override;

   private
    { Déclarations privées }
   
    // Modif BTP
//    CC : THValComboBox;
    // ------
    procedure MAJBaseTarif;
    end;

  Tof_AssistantTarif_mul = Class (TOF)
    procedure OnArgument (Arguments : String ) ; override;
    procedure AssistantTarifArt_mul;
   private
    { Déclarations privées }
    // Modif BTP
{$IFDEF BTP}
    CC : THValComboBox;
{$Endif}
    // ------

  end ;

const
// libellés des messages
TexteMessage: array[1..2] of string 	= (
          {1}  'Aucun élément sélectionné'
          {2} ,'La sélection est trop importante, vous devez la restreindre'
              );

implementation

Procedure Tof_TarifArt_Mul.OnArgument (Arguments : String ) ;
begin
inherited ;

GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GA_LIBREART', 10, '');

THValComboBox ( getcontrol('GA_TYPEARTICLE')).Plus := 'AND (CO_CODE="MAR" OR CO_CODE="PRE" OR CO_CODE="ARP")';
THEdit(getcontrol('XX_WHERE')).Text := 'AND (GA_TYPEARTICLE="MAR" OR GA_TYPEARTICLE="PRE" OR GA_TYPEARTICLE="ARP")';
SetControlVisible ('TGA_COLLECTION',False);
SetControlVisible ('GA_COLLECTION',False);

//uniquement en line
{*
TTabSheet(GetControl('PLIBRES')).tabvisible := false;
TFMul(ecran).SetDBListe ('BTTARIFART_S1');
*}

END;

procedure Tof_TarifArt_Mul.MAJBaseTarif;
var F : TFMul ;
    TOBArticle, TOBMul : TOB ;
    i_ind,iNbArt : integer ;
    Where, stSQL : string ;
    QQ : TQuery ;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
    begin
{$IFNDEF EAGLCLIENT}
    if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
{$ENDIF}
    HShowMessage('0;'+F.Caption+';'+TexteMessage[1]+';W;O;O;O;','','') ;
    exit;
    end;
Where := RecupWhereCritere(F.Pages) ;
stSQL := 'SELECT GA_ARTICLE,GA_CODEARTICLE,GA_PAHT,GA_PRHT,GA_DPA,GA_PMAP,GA_DPR,GA_PMRP,'+
         'GA_CALCPRIXHT,GA_CALCPRIXTTC,' +
{$IFDEF BTP}
         'GA_COEFFG,GA_DPRAUTO,GA_CALCPRIXPR,' +
{$ENDIF}
         'GA_CALCAUTOHT,GA_CALCAUTOTTC,GA_ARRONDIPRIX,GA_COEFCALCHT,GA_COEFCALCTTC,' +
         'GA_PVHT,GA_PVTTC  FROM ARTICLE ' + Where;
iNbArt := 0;
if F.FListe.AllSelected then
   begin
   QQ := OpenSQL('SELECT COUNT(GA_ARTICLE) NBART FROM ARTICLE ' + Where,true);
   if not QQ.Eof then iNbArt := QQ.FindField('NBART').AsInteger;
   Ferme(QQ);
   end else iNbArt := F.FListe.NbSelected;

if iNbArt> 3000 then
   begin
{$IFNDEF EAGLCLIENT}
   if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
{$ENDIF}
   HShowMessage('0;'+F.Caption+';'+TexteMessage[2]+';W;O;O;O;','','') ;
   exit;
   end;

TOBArticle := TOB.Create ('', Nil, -1) ;
if F.FListe.AllSelected then
   begin
   QQ := OpenSQL(stSQL,true);
   TOBArticle.LoadDetailDB('ARTICLE','','',QQ,False);
   Ferme(QQ);
   end else
   begin
   if   Where = '' then Where := ' WHERE GA_ARTICLE='
   else Where := ' AND GA_ARTICLE=';
   for i_ind:=0 to F.FListe.NbSelected-1 do
       begin
       TOBMul := TOB.Create ('ARTICLE', TobArticle, -1) ;
       F.FListe.GotoLeBOOKMARK(i_ind);
{$IFDEF EAGLCLIENT} // DBR Fiche 10095
       F.Q.TQ.Seek(F.FListe.Row-1) ;
{$ENDIF}
       QQ := OpenSQL(stSQL + Where +'"'+TFmul(Ecran).Q.FindField('GA_ARTICLE').AsString+'"',true);
       TOBMul.SelectDB('',QQ) ;
       Ferme(QQ);
       end;
   end;

if TOBArticle.Detail.Count <= 0 then exit ;
// Assistant de mise à jour
//uniquement en line
//BTAssist_MajTarifArt (TOBArticle);
Assist_MajTarifArt (TOBArticle) ;

if F.FListe.AllSelected then F.FListe.AllSelected:=False else F.FListe.ClearSelected;

F.bSelectAll.Down := False ;

TOBArticle.free ;

end;

//=========================================================================
//=========================================================================
Procedure TOF_AssistantTarif_mul.OnArgument (Arguments : String ) ;
var St : string;
    OkS3 : boolean ;
begin
inherited ;
St:=Arguments ;
// Affichage Domaine
{$IFDEF CCS3} OkS3:=True ; {$ELSE} OkS3:=False ; {$ENDIF}
if ((Not OkS3) and (THValComboBox(GetControl('TMPDOMAINE')).Values.Count>1)) then
   BEGIN
   SetControlVisible('GA_DOMAINE',True);
   SetControlVisible('TGA_DOMAINE',True);
   SetControlProperty('GA_DOMAINE','TabStop',True);
   END;
{$IFDEF BTP} 
CC := THValComboBox(ecran.FindComponent('GA_TYPEARTICLE'));
CC.Plus := 'AND CO_CODE <> "OUV" AND CO_CODE <> "POU" AND CO_CODE <> "EPO"';
{$ENDIF}
GCMAJChampLibre(TForm (Ecran),False,'COMBO','GA_LIBREART',10,'');
END;

procedure TOF_AssistantTarif_mul.AssistantTarifArt_mul;
var F : TFMul ;
    TOBArticle (*, TOBMul*) : TOB ;
    iInd, iNbArt, iPosGen : integer ;
    stSQL, stSqlCount, stWhere, stWhereCount, stChampArt : string ;
    QQ : TQuery ;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
    begin
{$IFNDEF EAGLCLIENT}
    if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
{$ENDIF}
    HShowMessage('0;'+F.Caption+';'+TexteMessage[1]+';W;O;O;O;','','') ;
    exit;
    end;

iNbArt := 0;
stWhere := RecupWhereCritere(F.Pages);
if (pos ('DIM', GetControlText('GA_STATUTART')) = 0) and
   (pos ('GEN', GetControlText('GA_STATUTART')) <> 0) then
    begin
    iPosGen := pos ('"GEN"', stWhere);
    stWhereCount := Copy (stWhere, 0, iPosGen) + 'DIM' +
                        Copy (stWhere, iPosGen + 4, Length (stWhere));
    stSqlCount := 'SELECT COUNT(GA_ARTICLE) NBART FROM ARTICLE ';
    if F.FListe.AllSelected then
        begin
        QQ := OpenSql (stSqlCount + stWhereCount, True);
        if not QQ.Eof then iNbArt := QQ.FindField('NBART').AsInteger;
        Ferme(QQ);
        end else
        begin
        if stWhereCount = '' then stWhereCount := ' WHERE (GA_CODEARTICLE="'
        else stWhereCount := stWhereCount + ' AND (GA_CODEARTICLE="';
        for iInd :=0 to F.FListe.NbSelected-1 do
            begin
            if iInd > 0 then stWhereCount := stWhereCount + ' OR GA_CODEARTICLE="';
            F.FListe.GotoLeBOOKMARK(iInd);
{$IFDEF EAGLCLIENT} // DBR Fiche 10095
            F.Q.TQ.Seek(F.FListe.Row-1) ;
{$ENDIF}
            stWhereCount := stWhereCount + TFmul(Ecran).Q.FindField('GA_CODEARTICLE').AsString +
                            '"';
            end;
        stWhereCount := stWhereCount + ')';
        QQ := OpenSql (stSqlCount + stWhereCount, True);
        if not QQ.Eof then iNbArt := QQ.FindField('NBART').AsInteger;
        Ferme(QQ);
        end;
    stWhere := stWhereCount;
    end else
    begin
    if (GetControlText('GA_STATUTART') = '') or
       ((pos ('DIM', GetControlText ('GA_STATUTART')) = 0) AND
        (pos ('GEN', GetControlText ('GA_STATUTART')) = 0) AND
        (pos ('UNI', GetControlText ('GA_STATUTART')) = 0)) then
        begin
        stWhere := stWhere + ' AND (GA_STATUTART="UNI" OR GA_STATUTART="DIM") ';
        end;
    iPosGen := pos ('"GEN"', stWhere);
    if iPosGen <> 0 then
        begin
        stWhereCount := Copy (stWhere, 0, iPosGen) + 'DIM' +
                            Copy (stWhere, iPosGen + 4, Length (stWhere));
        stWhere := stWhereCount;
        end;
    stChampArt := 'ARTICLE';
    if F.FListe.AllSelected then
        begin
        QQ := OpenSQL('SELECT COUNT(GA_ARTICLE) NBART FROM ARTICLE ' + stWhere, true);
        if not QQ.Eof then iNbArt := QQ.FindField('NBART').AsInteger;
        Ferme(QQ);
        end else
        begin
        iNbArt := F.FListe.NbSelected;
        if stWhere = '' then stWhere := ' WHERE (GA_' + stChampArt + '="'
        else stWhere := stWhere + ' AND (GA_' + stChampArt + '="';
        for iInd :=0 to F.FListe.NbSelected-1 do
            begin
            if iInd > 0 then stWhere  := stWhere  + ' OR GA_' + stChampArt + '="';
            F.FListe.GotoLeBOOKMARK(iInd);
{$IFDEF EAGLCLIENT} // DBR Fiche 10095
            F.Q.TQ.Seek(F.FListe.Row-1) ;
{$ENDIF}
            stWhere  := stWhere  + TFmul(Ecran).Q.FindField('GA_' + stChampArt).AsString +
                            '"';
            end;
        stWhere := stWhere + ')';
        end;
    end;

stSQL := 'SELECT GA_CODEARTICLE, GA_ARTICLE, GA_PAHT, GA_PVHT, GA_PVTTC, ' +
            'GA_STATUTART FROM ARTICLE ' + stWhere;

if iNbArt> 3000 then
   begin
{$IFNDEF EAGLCLIENT}
   if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
{$ENDIF}
   HShowMessage('0;'+F.Caption+';'+TexteMessage[2]+';W;O;O;O;','','') ;
   exit;
   end;

TOBArticle := TOB.Create ('', Nil, -1) ;
//if F.FListe.AllSelected then
   begin
   TOBArticle.LoadDetailFromSQL (stSQL);
   end;

if TOBArticle.Detail.Count > 0 then
    begin
    Assist_CreationTarifArt (TOBArticle);
    if F.FListe.AllSelected then F.FListe.AllSelected:=False else F.FListe.ClearSelected;
    F.bSelectAll.Down := False ;
    end;
TOBArticle.free ;
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLTarifArt_mul_Maj(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is Tof_TarifArt_Mul) then Tof_TarifArt_Mul(MaTOF).MAJBaseTarif else exit;
end;

procedure AGLAssistantTarif_mul (parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_AssistantTarif_mul) then TOF_AssistantTarif_mul(MaTOF).AssistantTarifArt_mul else exit;
end;

Initialization
registerclasses([Tof_TarifArt_Mul]);
registerclasses([Tof_AssistantTarif_mul]);
RegisterAglProc('TarifArt_mul_Maj',TRUE,1,AGLTarifArt_mul_Maj);
RegisterAglProc('AssistantTarif_mul',TRUE,1,AGLAssistantTarif_mul);
end.
