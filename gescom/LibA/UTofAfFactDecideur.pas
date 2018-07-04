unit UTofAfFactDecideur;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}
    eMul,
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db, Mul,
{$ENDIF}

      HCtrls,HEnt1,HMsgBox,UTOF, AffaireUtil,UTob,Grids,EntGC,FactTOB,
      Dicobtp,Saisutil,Factutil,M3FP,vierge,AFTableauBord,uEntCommun,UtilTOBPiece;


Type
     TOF_AFFactDecideur = Class (TOF)
     public
        procedure OnLoad ; override ;
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnClose ; override ;

     private
        zaff : string;
        DdateD,DdateF: TDateTime;
        znumech : integer;
        TobEch,TobAff : TOB;
        DEV : RDEVISE;
        zbonimali,zacpte,zprod,zprod_pres,zprod_four,zprod_frais : Double;
        ztot,zafact,zlig,zboni,zmali: Double;
        zprevu,zfact,zcout,zcoutpv: Double;

     procedure RechInfos;
     //procedure CalculTot;
     procedure AlimentationDetail(num : integer);
     //procedure TestEvent (Sender: TObject);
     Function  RemplirTOBAff ( CodeAffaire : String ;AcTobAff : TOB) : boolean ;
     procedure RecupActivite ( CodeAff,TypPres,NomZone : String ;DateD,DateF:TdateTime;var zcum : double) ;
     procedure AlimentationGlobal;
     procedure RecupPiece(CodeAff,NomZone,nat : String ; var zcum : double) ;
     procedure CalculTotalHTBase(CodeAff : String;zpou : double ; var zcum ,zesc,zrem: double) ;
      procedure Reaffichage(num : integer) ;
    END ;

 Type
     ParamTB = Class (TParamTBAffaire)
     END;

     const
	// libellés des messages de la TOF  AfFactDecideur
	TexteMsgAffaire: array[1..7] of string 	= (
          {1}        'La notion d''acompte est obligatoire'
          {2}        ,'Code Prestation invalide'
          {3}        ,'Date Invalide'
          {4}        ,'Dates Incompatibles'
          {5}        ,'Pas de données sélectionnées'
          {6}        ,'Génération en cours'
          {7}        ,'Vérifiez vos dates de génération'
                     );


implementation

{ TOF_AfAppreciation}


procedure TOF_AfFactDecideur.OnLoad;
Begin
 Inherited;

End;

procedure TOF_AfFactDecideur.OnClose;

Begin
 Inherited;
  TobEch.free;
  TobAff.free;
End;

procedure TOF_AfFactDecideur.OnArgument(stArgument : String );
Var //CC,CC1 : THEdit;
    Critere, Champ, valeur,ZdateD,ZdateF  : String;
    x : integer;


begin
Inherited;
// Recup des critères
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN                            
        X:=pos(':',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if Champ = 'AFA_AFFAIRE' then Zaff := Valeur;
        if Champ = 'AFA_NUMECHE'    then Znumech := StrToInt(Valeur);
        if Champ = 'ZDATEACT_D'    then ZDateD := Valeur;
        if Champ = 'ZDATEACT_F'    then ZDateF := Valeur;
        END;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;

 DDateD := StrToDate (ZDateD);
 DDateF := StrToDate (ZDateF);

SetControlText('ACT_DATEACTIVITE',zdateD);
SetControlText('ACT_DATEACTIVITE_',zdateF);
// ruse pour ne pas passer par le script
//CC := THEDIT(GetControl('ZBONI'));
//CC.OnExit:=TestEvent;

//CC1 := THEDIT(GetControl('ZMALI'));
//CC1.OnExit:=TestEvent;

ChargeCleAffaire(Nil,THEDIT(GetControl('AFA_AFFAIRE1')), THEDIT(GetControl('AFA_AFFAIRE2')),
THEDIT(GetControl('AFA_AFFAIRE3')),THEDIT(GetControl('AFA_AVENANT')), Nil, taConsult,zaff,False);

RechInfos;

AlimentationDetail(1);

AlimentationGlobal;
end;

{procedure TOF_AfFactDecideur.TestEvent (Sender: TObject);
var st : string;

begin

if ((Sender is THEdit) and (Thedit(Sender).Name='ZBONI') and  (Thedit(Sender).Modified)) then
  begin
    Thedit(Sender).Modified :=False ;
    st := Thedit(Sender).text;
    zboni := 0;
    if (st <> '') then
      Begin
      zboni := STRToFloat(st);
//         PGIBoxAf(TexteMsgAffaire[1],TitreHalley);
//         SetFocusControl('ZBONI');
//      exit;
      End;
    CalculTot;
 End;

 if ((Sender is THEdit) and (Thedit(Sender).Name='ZMALI') and  (Thedit(Sender).Modified)) then
  begin
    Thedit(Sender).Modified :=False ;
    st := Thedit(Sender).text;
    zmali := 0;
    if (st <> '') then
      Begin
      zmali := STRToFloat(st);
//         PGIBoxAf(TexteMsgAffaire[1],TitreHalley);
//         SetFocusControl('ZBONI');
//      exit;
      End;
    CalculTot;
 End;


end;
}

// Acces à l'échéance et à l'affaire , puis analyse des données
procedure TOF_AfFactDecideur.RechInfos;
var  LibDevise : string;

Begin

// zone saisisable
//   SetFocusControl('ZBONI');

// accès à l'affaire
   TOBAff:=TOB.Create('AFFAIRE',Nil,-1) ;
   RemplirTobAff(zaff,TobAff);
   SetControlText('TLIBAFF',TOBAff.GetValue('AFF_LIBELLE'));


// trt devise
   DEV.Code := TOBAff.GetValue('AFF_DEVISE');
   if (DEV.Code = '') then DEV.code := V_PGI.DevisePivot;
   GetInfosDevise(DEV) ;
   LibDevise := DEV.Libelle;

// Acces à l'échéance
  TOBEch:=TOB.Create('FACTAFF',Nil,-1) ;
  RemplirTobEcheance(zaff,znumech,TobEch);
   // mcd 31/05/02 si tob vode, aff pas renseigné ... SetControlText('AFA_AFFAIRE',TOBEch.GetValue('AFA_AFFAIRE'));
  SetControlText('AFA_AFFAIRE',zaff);
  SetControlText('AFA_DATEECHE',TOBEch.GetValue('AFA_DATEECHE'));
  SetControlText('AFA_REPRISEACTIV',TOBEch.GetValue('AFA_REPRISEACTIV'));

end;


procedure TOF_AfFactDecideur.Reaffichage(num : integer);
Begin
  TobEch.free; TobEch:=nil;
  TobAff.free; TobAff:=nil;
  RechInfos ;
  Alimentationdetail(num);
  AlimentationGlobal;
End;

// Alimentation des données propres à l'échéance
// num = 0 1er passage  <>0 reraffichage
procedure TOF_AfFactDecideur.AlimentationDetail(num :integer);
var sboni, smali,zdat,zlib,TypAct,zcombo,TypGen,zst : String;
    T : Tgroupbox;
    toppre,topfou,topfra : boolean;
    zpou,zesc,zrem : Double;

begin

if (TOBEch=nil) or (TOBAff=nil) then exit;

    TypAct := TOBEch.GetValue('AFA_REPRISEACTIV');
    TypGen := TOBAff.GetValue('AFF_GENERAUTO');
// Libelle Groupe Detail
    zdat := DateToStr(TOBEch.GetValue('AFA_DATEECHE'));
    if (TOBEch.GetValue('AFA_LIQUIDATIVE') = 'X') then
      Begin
        zlib := Format('Facture Liquidative du %s',[zdat]);
      End
    else
      Begin
        zlib := Format('Echéance du %s',[zdat]);
      End;

// blocage si Ech facturée
    if (TOBEch.GetValue('AFA_ECHEFACT') = 'X') then
    Begin
      SetControlVisible('ZLIBETAT',true);
 //     SetControlEnabled('ZBONI',false);
 //     SetControlEnabled('ZMALI',false);
      zlib := zlib + ' (facturée)';
    End
    else
    Begin
//      SetControlEnabled('ZBONI',false);
//      SetControlEnabled('ZMALI',false);
    End;

    T:=TGroupBox(GetControl('FE_APPREC'));
    T.caption := zlib;

// production
  zprod :=0;
  zprod_pres := 0;
  zprod_four := 0;
  zprod_frais := 0;
  zacpte:=0;
  zlig:=0;
  ztot:=0;
  zbonimali:=0;
  zmali:=0;
  zafact:=0;
  zesc := 0.0 ; zrem := 0.0;


// Prod Prestations   Frais  Fourniture


  // alimentation du mul_combo type ligne activite en fct de l'activite reprise
  toppre := false; topfra := false; topfou := false;
  zcombo := '';
  Trt_Activiterepris(typact,zcombo,toppre,topfou,topfra);

  SetControlEnabled('ZPROD_PRES',true);
  SetControlEnabled('TZPROD_PRES',true);
  SetControlEnabled('ZPROD_FRAIS',true);
  SetControlEnabled('TZPROD_FRAIS',true);
  SetControlEnabled('ZPROD_FOUR',true);
  SetControlEnabled('TZPROD_FOUR',true);
  SetControlEnabled('ZPROD',true);
  SetControlEnabled('TZPROD',true);
  SetControlEnabled('BPROD',true);

  if not(toppre) then
    Begin
      SetControlEnabled('ZPROD_PRES',false);
      SetControlEnabled('TZPROD_PRES',false);
    End;
  if not(topfra) then
    Begin
      SetControlEnabled('ZPROD_FRAIS',false);
      SetControlEnabled('TZPROD_FRAIS',false);
    End;
  if not(topfou) then
    Begin
      SetControlEnabled('ZPROD_FOUR',false);
      SetControlEnabled('TZPROD_FOUR',false);
    End;
  if (not(toppre) and not(topfra) and not(topfou) and (typgen <> 'ACT'))then
  Begin
      SetControlEnabled('ZPROD',false);
      SetControlEnabled('TZPROD',false);
      SetControlEnabled('BPROD',false);
  End;

  if (toppre) then
    RecupActivite ( zaff,'PRE','ACT_TOTVENTE',DDateD,DDateF,zprod_pres) ;
  if (topfou) then
    RecupActivite ( zaff,'MAR','ACT_TOTVENTE',DDateD,DDateF,zprod_four) ;
  if (topfra ) then
    RecupActivite ( zaff,'FRA','ACT_TOTVENTE',DDateD,DDateF,zprod_frais) ;

  SetControlText('ZPROD_PRES',StrF00(zprod_pres,DEV.Decimale));
  SetControlText('ZPROD_FOUR',StrF00(zprod_four,DEV.Decimale));
  SetControlText('ZPROD_FRAIS',StrF00(zprod_frais,DEV.Decimale));
// Prod frais

// Prod Fournitures

// Prod Global
  zprod := zprod_pres + zprod_four + zprod_frais;
  SetControlText('ZPROD',StrF00(zprod,DEV.Decimale));


// Forfait
  if (typgen = 'FOR') then
  Begin
    SetControlEnabled('TZACPTE',true);
    SetControlEnabled('ZACPTE',true);
    zacpte := TOBEch.GetValue('AFA_MONTANTECHEDEV');
    SetControlText('ZACPTE',StrF00(zacpte,DEV.Decimale));
  End
  else
  Begin
    SetControlEnabled('TZACPTE',false);
    SetControlEnabled('ZACPTE',false);
  End;


// Lignes

  if (copy(typgen,1,1) = 'P') then
  Begin
    zpou :=  TOBEch.GetValue('AFA_POURCENTAGE');
    CalculTotalHTBase(zaff,zpou,zlig,zesc,zrem);

    SetControlEnabled('TZLIG',true);
    SetControlEnabled('ZLIG',true);
    SetControlText('ZLIG',StrF00(zlig,DEV.Decimale));
  End
  Else
  Begin
    SetControlEnabled('TZLIG',false);
    SetControlEnabled('ZLIG',false);
  End;

// Facturable
  ztot := zprod + zacpte + zlig;
  SetControlText('ZTOT',StrF00(ztot,DEV.Decimale));

// boni / Mali
  if (num = 0) then
  Begin
    zbonimali :=  TobEch.getvalue('AFA_BONIMALI');
    if (zbonimali >= 0)  then
      Begin
        zboni := zbonimali;
        sboni := StrF00(zbonimali,DEV.Decimale)
      End
    else
      Begin
        zmali := zbonimali;
        zbonimali := zbonimali * (-1);
        smali := StrF00(zbonimali,DEV.Decimale);
      End;

//    SetControlText('ZBONI',sboni);
//    SetControlText('ZMALI',smali);
  End;

// A facturer
  zafact := zprod + zacpte + zlig + zboni - zmali;
  // PL le 21/10/03 : on controle si ce champ existe car pas toujours le cas
  if (GetControl ('ZAFACT') <> nil) then
    SetControlText('ZAFACT',StrF00(zafact,DEV.Decimale));
    
//commentaire Etoile
  if ((zrem <> 0) or (zesc <> 0)) then
    zst := format('(*)Hors remise %4.2n et escompte %4.2n',[zrem,zesc])
  else
    zst := '(*)Hors remise et escompte ';

  SetCOntrolText('TETOILE',zst);
end;

// Alimentation des données propres à l'échéance
procedure TOF_AfFactDecideur.AlimentationGlobal;

begin

   
// Prevu global Affaire
if (TOBAff<>nil) then
   zprevu := TobAff.GetValue('AFF_TOTALHTGLODEV');

   SetControlText('ZPREVU',StrF00(zprevu,DEV.Decimale));

// Facturé
  zfact := 0;
    ///mcd 31/05/02 RecupPiece ( zaff,'GP_TOTALHTDEV','FAC',zfact) ;
  RecupPiece ( zaff,'GL_TOTALHTDEV','FAC;AVC;FRE',zfact) ;
  SetControlText('ZFACT',StrF00(zfact,DEV.Decimale));

// Cout PR
  zcout := 0;
  RecupActivite ( zaff,'TOU','ACT_TOTPRCHARGE',idate1900,DDateF,zcout) ;
  SetControlText('ZCOUT',StrF00(zcout,DEV.Decimale));

  zcoutpv := 0;
  RecupActivite ( zaff,'TOU','ACT_TOTVENTE',idate1900,DDateF,zcoutpv) ;
  SetControlText('ZCOUTPV',StrF00(zcoutpv,DEV.Decimale));

  // resul / prevu
  SetControlText('ZRESPREV',StrF00(zfact-zprevu,DEV.Decimale));
  SetControlText('ZRESCOUT',StrF00(zfact-zcoutpv,DEV.Decimale));
end;

(* supprimer suite conseil , car non utilisé, voir pourquoi ?
procedure TOF_AfFactDecideur.CalculTot;
BEGIN
// Facturable
  ztot := zprod + zacpte + zlig;
  SetControlText('ZTOT',StrF00(ztot,DEV.Decimale));

// A facturer
//  zafact := zprod + zacpte + zlig + zboni - zmali;
//  SetControlText('ZAFACT',StrF00(zafact,DEV.Decimale));

END;
*)


Function TOF_AfFactDecideur.RemplirTOBAff ( CodeAffaire : String ; AcTobAff : TOB) : boolean ;
Var Q : TQuery ;
BEGIN
Result:=True ;
if CodeAffaire='' then  AcTOBAff.InitValeurs(False)
else if CodeAffaire<>AcTOBAff.GetValue('AFF_AFFAIRE') then
   BEGIN
        // fait pour une affaire seuelemnt. on laise....
   Q:=OpenSQL('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+CodeAffaire+'"',True) ;
   If (Not Q.EOF) then
   Result:=AcTOBAff.SelectDB('',Q)
   Else Result:=False;
   Ferme(Q);
   END ;
END ;

procedure TOF_AfFactDecideur.RecupActivite(CodeAff,TypPres,NomZone : String ;DateD,DateF:TdateTime;var zcum : double) ;
Var req,whereplus,wheredate ,wherefac: string;
    QQ : tquery;
BEGIN
  wheredate := ' AND ACT_DATEACTIVITE <= "'+ usdatetime(DateF)+'"';
  wheredate := wheredate + ' AND ACT_DATEACTIVITE >= "'+ usdatetime(DateD)+'"';

  if (TypPres <> 'TOU') then
  Begin
    whereplus := ' and ACT_TYPEARTICLE = "'+TypPres+'"';
    wherefac := ' and ACT_ACTIVITEREPRIS="F" ';
  End
  else
  Begin
    whereplus := '';
    wherefac := ' and ACT_ACTIVITEREPRIS <> "N" ';
  End;
  Req := 'SELECT SUM('+NomZone+') FROM ACTIVITE WHERE ACT_AFFAIRE = "'+
  CodeAff+'"'+whereplus+wheredate+wherefac;
  QQ:=OpenSQL(Req,TRUE);
  if Not QQ.EOF then
    zcum:=QQ.Fields[0].AsFloat;

  Ferme(QQ)
END;

procedure TOF_AfFactDecideur.RecupPiece(CodeAff,NomZone,nat : String ; var zcum : double) ;
Var req ,reqnat,critere: string;
    QQ : tquery;
    ii: integer;
BEGIN
// mcd 31/05/02 on part des lignes et on prend toutes les natures ventes ..
//  Req := 'SELECT SUM('+NomZone+') FROM PIECE WHERE GP_AFFAIRE = "'+
//  CodeAff+'" AND GP_NATUREPIECEG = "'+nat+'"';
Critere:=(Trim(ReadTokenSt(nat)));
reqNat:=' AND (';
ii:=0;
While (Critere <>'') do
    BEGIN
    if ii <>0 then Reqnat:=ReqNat + ' OR '
       else ii:=1;
    ReqNat:=reqNat +  ' GL_NATUREPIECEG="'+critere+'"';
    Critere:=(Trim(ReadTokenSt(nat)));
    END;
ReqNat:=ReqNat + ' )';

Req := 'SELECT SUM('+NomZone+') FROM LIGNE WHERE GL_AFFAIRE = "'+
      CodeAff+'" '+Reqnat+' AND GL_TYPELIGNE="ART"';
  QQ:=OpenSQL(Req,TRUE);
  if Not QQ.EOF then
    zcum:=QQ.Fields[0].AsFloat;

  Ferme(QQ)
END;

// calcul du Total HT de l'affaire AVANT Remise et Escompte
procedure TOF_AfFactDecideur.CalculTotalHTBase(CodeAff : String;zpou : double ; var zcum,zesc,zrem : double) ;
Var req : string;
    QQ : tquery;
    CleDoc : R_Cledoc;

BEGIN
  zcum := 0.0;
  SelectPieceAffaire(CodeAff, 'AFF', CleDoc);
  req := 'SELECT sum(GL_MONTANTHTDEV) FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,False)+ 'and GL_TYPELIGNE="ART"';

  QQ:=OpenSQL(Req,TRUE);
  if Not QQ.EOF then
    zcum:=QQ.Fields[0].AsFloat;

  if (zcum <> 0) then
    zcum :=  Arrondi(zcum*zpou/100.0,V_PGI.OkDecV) ;
  Ferme(QQ);

  req := 'SELECT GP_REMISEPIED,GP_ESCOMPTE FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False);
  QQ:=OpenSQL(Req,TRUE);
  if Not QQ.EOF then
  Begin
    zrem:=QQ.Fields[0].AsFloat;
    zesc:=QQ.Fields[1].AsFloat;
  End;

// if (zcum <> 0) then
//    zcum :=  Arrondi(zcum*zpou/100.0,2) ;
  Ferme(QQ);

END;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLReaffichage(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;

if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;

if (MaTOF is TOF_AfFactDecideur) then TOF_AfFactDecideur(MaTOF).Reaffichage(1) else exit;
end;

Initialization
registerclasses([TOF_AfFactDecideur]);
RegisterAglProc('Reaffichage',TRUE,1,AGLReaffichage);
end.
