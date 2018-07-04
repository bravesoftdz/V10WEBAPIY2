unit UTofTarif;

interface

uses  StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
      eMul,Maineagl,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,DBGrids,HDB,Fiche,Fe_Main,
{$ENDIF}
      forms,sysutils,ComCtrls,AssistDupTarif,
      HCtrls,HEnt1,HMsgBox,UTOF, HDimension,UTOM,AGLInit,
      Utob,Messages,HStatus,M3VM , M3FP, Hqry,
      AssistMajTarif;

Type

     TOF_TarifMaj_Mul = Class (TOF)

        procedure OnArgument (Arguments : String ) ; override ;
//        procedure OnLoad ; override ;
//        procedure OnUpdate ; override ;
        procedure MajTarif;
        procedure DupTarif;
        procedure SupTarif;
        procedure SuppressionTarif;
     END ;

     TOF_TarifCon_Mul = Class (TOF)
        procedure OnArgument (Arguments : String ) ; override ;
     END ;

var
   TOBTarif : TOB;

procedure AGLTarifMaj_mul_Maj(parms:array of variant; nb: integer ) ;
          // mcd 12/06/02
Procedure AFLanceFiche_Mul_Maj_Tarif(Range,Argument:string);
Procedure AFLanceFiche_Mul_Maj_TarifFou;
Procedure AFLanceFiche_Mul_Consult_TarifFou;

const
// libellés des messages
TexteMessage: array[1..2] of string 	= (
          {1}  'Aucun élément sélectionné'
          {2} ,'Confirmez vous la suppression ?'
              );

implementation

Procedure TOF_TarifCon_Mul.OnArgument (Arguments : String ) ;
begin
inherited ;
//SetControlText('GF_REGIMEPRIX','GLO');
// Modif BTP
SetControlProperty ('GF_ARTICLE','plus','AND GA_TYPEARTICLE <> "POU" AND GA_TYPEARTICLE <> "OUV" AND GA_TYPEARTICLE <> "EPO"');
//
end;

Procedure TOF_TarifMaj_Mul.OnArgument (Arguments : String ) ;
var F : TFMul ;
begin
F:=TFMul(Ecran);
// Modif BTP
SetControlProperty ('GF_ARTICLE','plus','AND GA_TYPEARTICLE <> "POU" AND GA_TYPEARTICLE <> "OUV" AND GA_TYPEARTICLE <> "EPO"');
//
// FICHE_10314 DEBUT
//if Arguments = 'DUP' then
if Pos ('DUP', Arguments) > 0 then
// FICHE_10314 FIN
    begin
    F.Caption := 'Duplication des tarifs';
    UpdateCaption(Ecran);
    end
// FICHE_10314 DEBUT
//    else if Arguments = 'SUP' then
    else if Pos ('SUP', Arguments) > 0 then
// FICHE_10314 FIN
        begin
// FICHE_10314 DEBUT
        // F.Caption := 'Suppression des tarifs';
        if pos ('VEN', Arguments) > 0 then
        begin
          SetControlText ('NATUREAUXI', 'CLI');
          F.Caption := 'Suppression des tarifs clients';
        end else
        begin
          SetControlText ('NATUREAUXI', 'FOU');
          F.Caption := 'Suppression des tarifs Fournisseurs';
          SetControlText ('XX_WHERENAT', 'GF_NATUREAUXI="FOU"');
          SetControlProperty ('GF_TIERS', 'DataType', 'GCTIERSFOURN');
        end;
// FICHE_10314 FIN
        UpdateCaption(Ecran);
        // DBR : Vu que ce ne sont pas les champs GF_DATEFIN et GF_DATEDEBUT qui sont ceux de la saisie
        // faut faire sur les champs DATE et DATE_ : vu avec fiche 10314
//        SetControlVisible('GF_DATEFIN', False); est déja invisible
        SetControlVisible('TGF_DATEFIN', False); // Gardé
        SetControlVisible ('DATE_', False); // Ajouté
        SetControlText('TGF_DATEDEBUT','Date inférieure à'); // Gardé
//        SetControlProperty('GF_DATEDEBUT','Comparaison','Superieur'); C'est pas ca par rapport au libellé affiché
          // De plus l'operteur 'Comparaison' n'existe pas !
        THEdit(GetControl('GF_DATEDEBUT')).Operateur := Inferieur; // Ajouté
//        SetControlText('GF_DATEDEBUT',DateToStr(Date)); DATE est le champ de saisie et affecte apres GF_DATEDEBUT
        SetControlText('DATE', DateToStr(Date)); // Ajouté
        end;
//SetControlText('GF_REGIMEPRIX','GLO');
inherited ;
end;

procedure TOF_TarifMaj_Mul.MajTarif ;
var F : TFMul ;
    TOBMul, TOBMD, TOBTD : TOB ;
    i_ind : integer ;
    Where : string ;
    QQ : TQuery ;
BEGIN
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
    BEGIN
    //MessageAlerte('Aucun élément sélectionné');
{$IFDEF EAGLCLIENT}
{$ELSE}
    if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
{$ENDIF}
    HShowMessage('0;'+F.Caption+';'+TexteMessage[1]+';W;O;O;O;','','') ;
    exit;
    END;
Where := RecupWhereCritere(F.Pages) ;
TobTarif := TOB.Create ('', Nil, -1) ;
QQ := OpenSQL('SELECT * FROM TARIF ' + Where, True) ;
if not QQ.EOF then TOBTarif.LoadDetailDB('TARIF', '', '', QQ, FAlse) else TOBTarif := nil;
Ferme (QQ) ;
if TOBTarif = nil then exit ;
if Not F.FListe.AllSelected then
    BEGIN
    TOBMul := TOB.Create ('MajTarif', Nil, -1) ;
    for i_ind:=0 to F.FListe.NbSelected-1 do
        BEGIN
        TOBMD := TOB.Create ('', TOBMul, -1);
        TOBMD.AddChampSup('GF_TARIF', True) ;
        F.FListe.GotoLeBOOKMARK(i_ind);
        TOBMD.PutValue ('GF_TARIF', TFmul(Ecran).Q.FindField('GF_TARIF').asstring) ;
        END;
    TOBMul.Detail.Sort('GF_TARIF') ;
    For i_ind := TOBTarif.Detail.Count-1 downto 0 do
        BEGIN
        TOBTD := TOBTarif.Detail[i_ind] ;
        TOBMD := TOBMul.FindFirst(['GF_TARIF'], [TOBTD.GetValue ('GF_TARIF')], False) ;
        if TOBMD=Nil then TOBTD.Free;
        END;
    TOBMul.free ;
    END;
// Assistant de mise à jour
Assist_MajTarif (TOBTarif) ;
if F.FListe.AllSelected then F.FListe.AllSelected:=False else F.FListe.ClearSelected;
F.bSelectAll.Down := False ;
TOBTarif.free ;
END;

procedure TOF_TarifMaj_Mul.DupTarif;
var F : TFMul ;
    TOBMul, TOBMD, TOBTD : TOB ;
    i_ind : integer ;
    Where : string ;
    QQ : TQuery ;
BEGIN
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
    BEGIN
    //MessageAlerte('Aucun élément sélectionné');
{$IFDEF EAGLCLIENT}
{$ELSE}
    if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
{$ENDIF}
    HShowMessage('0;'+F.Caption+';'+TexteMessage[1]+';W;O;O;O;','','') ;
    exit;
    END;
Where := RecupWhereCritere(F.Pages) ;
TobTarif := TOB.Create ('', Nil, -1) ;
QQ := OpenSQL('SELECT * FROM TARIF ' + Where, True) ;
if not QQ.EOF then TOBTarif.LoadDetailDB('TARIF', '', '', QQ, FAlse) else TOBTarif := nil;
Ferme (QQ) ;
if TOBTarif = nil then exit ;
if Not F.FListe.AllSelected then
    BEGIN
    TOBMul := TOB.Create ('MajTarif', Nil, -1) ;
    for i_ind:=0 to F.FListe.NbSelected-1 do
        BEGIN
        TOBMD := TOB.Create ('', TOBMul, -1);
        TOBMD.AddChampSup('GF_TARIF', True) ;
        F.FListe.GotoLeBOOKMARK(i_ind);
{$IFDEF EAGLCLIENT}
        F.Q.TQ.Seek(F.FListe.Row-1) ;
{$ENDIF}
        TOBMD.PutValue ('GF_TARIF', TFmul(Ecran).Q.FindField('GF_TARIF').asstring) ;
        END;
    TOBMul.Detail.Sort('GF_TARIF') ;
    For i_ind := TOBTarif.Detail.Count-1 downto 0 do
        BEGIN
        TOBTD := TOBTarif.Detail[i_ind] ;
        TOBMD := TOBMul.FindFirst(['GF_TARIF'], [TOBTD.GetValue ('GF_TARIF')], False) ;
        if TOBMD=Nil then TOBTD.Free;
        END;
    TOBMul.free ;
    END;
// Assistant de mise à jour

Assist_DupTarif (TOBTarif) ;

if F.FListe.AllSelected then F.FListe.AllSelected:=False else F.FListe.ClearSelected;
F.bSelectAll.Down := False ;
TOBTarif.free ;
END;

procedure TOF_TarifMaj_Mul.SupTarif;
var F : TFMul ;
    ioerr : TIOErr ;
    TOBJnal : TOB ;
    TOBMul, TOBMD, TOBTD : TOB ;
    i_ind, NumEvt : integer ;
    Where : string ;
    QQ : TQuery ;
    ListRecap : TStrings;
BEGIN
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
    BEGIN
    //MessageAlerte('Aucun élément sélectionné');
{$IFDEF EAGLCLIENT}
{$ELSE}
    if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
{$ENDIF}
    HShowMessage('0;'+F.Caption+';'+TexteMessage[1]+';W;O;O;O;','','') ;
    exit;
    END
    else
    begin
    if HShowMessage('1;'+F.Caption+';'+TexteMessage[2]+';Q;YNC;Y;','','') <> mrYes then Exit;
    end;
Where := RecupWhereCritere(F.Pages) ;
TobTarif := TOB.Create ('', Nil, -1) ;
QQ := OpenSQL('SELECT * FROM TARIF ' + Where, True) ;
if not QQ.EOF then TOBTarif.LoadDetailDB('TARIF', '', '', QQ, FAlse) else TOBTarif := nil;
Ferme (QQ) ;
if TOBTarif = nil then exit ;
if Not F.FListe.AllSelected then
    BEGIN
    TOBMul := TOB.Create ('SupTarif', Nil, -1) ;
    for i_ind:=0 to F.FListe.NbSelected-1 do
        BEGIN
        TOBMD := TOB.Create ('', TOBMul, -1);
        TOBMD.AddChampSup('GF_TARIF', True) ;
        F.FListe.GotoLeBOOKMARK(i_ind);
{$IFDEF EAGLCLIENT}
        F.Q.TQ.Seek(F.FListe.Row-1) ;
{$ENDIF}
        TOBMD.PutValue ('GF_TARIF', TFmul(Ecran).Q.FindField('GF_TARIF').asstring) ;
        END;
    TOBMul.Detail.Sort('GF_TARIF') ;
    For i_ind := TOBTarif.Detail.Count-1 downto 0 do
        BEGIN
        TOBTD := TOBTarif.Detail[i_ind] ;
        TOBMD := TOBMul.FindFirst(['GF_TARIF'], [TOBTD.GetValue ('GF_TARIF')], False) ;
        if TOBMD=Nil then TOBTD.Free;
        END;
    TOBMul.free ;
    END;

NumEvt:=0 ;
ListRecap := TStringList.Create;
ListRecap.Text := '';
TOBJnal:=TOB.Create('JNALEVENT', Nil, -1) ;
TOBJnal.PutValue('GEV_TYPEEVENT', 'TAR');
TOBJnal.PutValue('GEV_LIBELLE', F.Caption);
TOBJnal.PutValue('GEV_DATEEVENT', Date);
TOBJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
ioerr := Transactions(SuppressionTarif,2);
QQ:=OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',True) ;
if Not QQ.EOF then NumEvt:=QQ.Fields[0].AsInteger ;
Ferme(QQ) ;
Inc(NumEvt) ;
TOBJnal.PutValue('GEV_NUMEVENT', NumEvt);

Case ioerr of
        oeOk  : BEGIN
                TOBJnal.PutValue('GEV_ETATEVENT', 'OK');
                ListRecap.Add('');
                ListRecap.Add(IntToStr(TOBTarif.detail.count) + ' enregistrement(s) supprimé(s)');
                TOBJnal.PutValue('GEV_BLOCNOTE', ListRecap.Text);
                END;
    oeUnknown : BEGIN
                MessageAlerte('ATTENTION : Tarif non supprimé !') ;
                ListRecap.Add('');
                ListRecap.Add('ATTENTION : Tarif non supprimé !');
                TOBJnal.PutValue('GEV_ETATEVENT', 'ERR');
                END ;
    oeSaisie  : BEGIN
                MessageAlerte('ATTENTION : Ce tarif, en cours de traitement par un autre utilisateur, n''a pas été supprimé !') ;
                ListRecap.Add('');
                ListRecap.Add('ATTENTION : Ce tarif, en cours de traitement par un autre utilisateur, n''a pas été supprimé !');
                TOBJnal.PutValue('GEV_ETATEVENT', 'ERR');
                END ;
   END ;
TOBJnal.InsertDB(Nil) ;
TOBJnal.Free ;

if F.FListe.AllSelected then F.FListe.AllSelected:=False else F.FListe.ClearSelected;
F.bSelectAll.Down := False ;
TOBTarif.free ;
END;

procedure TOF_TarifMaj_Mul.SuppressionTarif;
begin
TOBTarif.DeleteDB;
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLTarifMaj_mul_Maj(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_TarifMaj_Mul) then TOF_TarifMaj_Mul(MaTOF).MajTarif else exit;
end;

procedure AGLTarifDup_mul_Maj(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
     st1, st2 : string;
begin
F:=TForm(Longint(Parms[0])) ;
st1 := string(Parms[1]);
st2 := string(Parms[2]);
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_TarifMaj_Mul) then TOF_TarifMaj_Mul(MaTOF).DupTarif else exit;
end;

procedure AGLTarifSup_mul_Maj(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
     st1, st2 : string;
begin
F:=TForm(Longint(Parms[0])) ;
st1 := string(Parms[1]);
st2 := string(Parms[2]);
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_TarifMaj_Mul) then TOF_TarifMaj_Mul(MaTOF).SupTarif else exit;
end;

Procedure AFLanceFiche_Mul_Maj_Tarif(Range,Argument:string);
begin
AGLLanceFiche ('AFF','AFTARIFMAJ_MUL',range,'',Argument);
end;
Procedure AFLanceFiche_Mul_Maj_TarifFou;
begin
AGLLanceFiche ('AFF','AFTARIFFOUMAJ_MUL','','','');
end;
Procedure AFLanceFiche_Mul_Consult_TarifFou;
begin
AGLLanceFiche ('AFF','AFTARIFFOUCON_MUL','','','');
end;


Initialization
registerclasses([TOF_TarifMaj_Mul]);
registerclasses([TOF_TarifCon_Mul]);
RegisterAglProc('TarifMaj_mul_Maj',TRUE,1,AGLTarifMaj_mul_Maj);
RegisterAglProc('TarifDup_mul_Maj',TRUE,1,AGLTarifDup_mul_Maj);
RegisterAglProc('TarifSup_mul_Maj',TRUE,1,AGLTarifSup_mul_Maj);
end.
