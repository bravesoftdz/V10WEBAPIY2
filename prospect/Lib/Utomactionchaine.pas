unit Utomactionchaine;

interface

uses Classes,Controls,forms,sysutils,stdctrls,
     HCtrls,HEnt1,HMsgBox,UTOM, UTob,M3FP,TiersUtil,Paramsoc, MailOl,
{$IFDEF EAGLCLIENT}
      MaineAGL,eFiche,UtileAGL,
{$ELSE}
      db,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}DBCtrls,
      Fe_Main,HsysMenu,Fiche,
{$ENDIF}
{$IFDEF POURVOIR}
OutlookView_TLB,
{$ENDIF}

Graphics,grids,windows,EntRT,UtilRT,AglInit,
EntGC, UtilPGI,menus
{$IFDEF AFFAIRE}
        ,AffaireUtil
{$ENDIF}

{$IFDEF VER150}
    ,Variants
{$ENDIF}

 ,FactUtil,HTB97, wCommuns, UtilChainage, UtilAction,uEntCommun
 ;

Type
 TOM_ACTIONSCHAINEES = Class (TOM)
   Public
   stTiers,stDerDate,stTermine,stNatureauxi :string;
   Creation,OrigineIntervention:boolean;
   NumChainage : Integer;
   procedure OnNewRecord  ; override ;
   procedure OnUpdateRecord  ; override ;
   procedure OnAfterUpdateRecord  ; override ;
   procedure OnArgument (Arguments : String ); override ;
   procedure OnDeleteRecord ; override ;
   procedure OnLoadRecord ; override ;
   procedure OnLoadAlerte; override ;
   procedure OnChangeField (F : TField) ; override ;
   procedure OnClose ; override ;
   procedure OnClickTermine(Sender: TObject);
   procedure EnvoiMessage (ListeResp,MentionBlc,MailAuto:String; Termine : boolean; St : HTStringList; PiecesVisees : String);
   function GetInfoGed : string ;
   private
   LesColonnesPiece : string ;
   GP : THGRID ;
   stProduitpgi : string;
   soRtchrappel: boolean;
   soRtmesslibact: boolean;
   soRtmessbl: boolean;
   soRtchchoixrappel:string;
   soRtgestinfos001: boolean;
   soRtmesstypeact: boolean;
   Generation_Tiers,Generation_VenAch,Generation_Chainage,Generation_Affaire,Generation_Ressource,
   Generation_LibTiers :String;
   TobActions,TobPiece : tob;
   procedure DessineCellGP ( ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
   procedure GPLigneDClick (Sender: TObject);
   procedure RTChainageAttacheAction;
   procedure ChainageNouvellePiece;
   procedure RTGenereLotActions;
   procedure ChargeGridPieces ;
   procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
   procedure RappelClick(Sender: TObject);
   procedure CreationActionsChainages;
   procedure TOBCopieChamp(FromTOB, ToTOB : TOB);
   procedure ActEnvoiMessage (TypeMessage:String;TobActionSav:tob;MailAuto:String);
   Procedure CreatInfosCompl(TobActionsChainage : Tob; StAuxi : String; INumAct : Integer);
{$IFDEF AFFAIRE}
    procedure RechercheAffaire (Sender: TObject);
    procedure EffaceAffaire (Sender: TObject);
    procedure AffaireEnabled( grise : boolean);
{$ENDIF}
    procedure RTAppelGED_OnClick (Sender: TObject) ;
    procedure GestionBoutonGED;
    procedure ListAlerte_OnClick_rch(Sender: TObject);
    procedure Alerte_OnClick_rch(Sender: TObject);
 END ;
procedure AGLRTChainageAttacheAction( parms: array of variant; nb: integer ) ;
procedure AGLRTChainagenouvellePiece( parms: array of variant; nb: integer )  ;
procedure AGLRTGenereLotActions( parms: array of variant; nb: integer )  ;
function Chainage_MyAfterImport (Sender : TObject): string ;
procedure Chainage_GestionBoutonGed (Sender : TObject) ;

procedure DessineCell (GG : Thgrid; ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
const
  ANNULEE  : String = 'ANU';
  REALISEE : String = 'REA';
  RESPONSABLE : string = 'RSP';
  INTERLOCUTEUR : string = 'INT';

  // libellés des messages
  TexteMessage: array[1..8] of string 	= (
    {1} 'Le libellé du chainage d''actions doit être renseigné'
    {2} ,'Vous ne pouvez pas effacer ce chaînage, il est actuellement utilisé dans les actions'
    {3} ,'Le code tiers n''existe pas'
    {4} ,'Type de chaînage inexistant'
    {5} ,'Le code tiers doit être renseigné'
    {6} ,'Code affaire incorrect'
    {7} ,'Chaînage terminé, pièce(s) visée(s) : '
    {8} ,'La date fin doit être supérieure ou égale à la date début'    
    );

implementation
uses CalcOleGenericAff,Facture
{$IFNDEF EAGLSERVER}
  ,UtilAlertes,YAlertesConst,EntPgi
{$ENDIF EAGLSERVER}
;

procedure TOM_ACTIONSCHAINEES.OnArgument(Arguments: String);
var Critere,ChampMul,ValMul : string;
    x : integer;
    Tcb,Termine: TCheckBox;
{$IFDEF AFFAIRE}
    bAff : TToolBarButton97;
{$ENDIF}
begin
  inherited;
  OrigineIntervention:=false;
  if pos ('ORIGINEINTERVENTION',Arguments) <> 0 then
     OrigineIntervention:=true;

  TobActions := nil;
  TobPiece := nil;
  stProduitpgi:='';
  stTermine:='';
  stTiers := ''; stDerDate:=DateToStr(iDate1900);
  { éléments passés à partir de RTGenereChainage }
  Generation_Tiers:='';
  Generation_VenAch:='';
  Generation_Chainage:='';
  Generation_Affaire:='';
  Generation_Ressource:='';
  Generation_LibTiers:='';
  Generation_Tiers  := GetArgumentString(Arguments,'GENETIERS');
  if Generation_Tiers <> '' then
    begin
    Generation_LibTiers  := GetArgumentString(Arguments,'GENELIBELLE',false);
    Generation_VenAch  := GetArgumentString(Arguments,'GENEVENACH');
    Generation_Chainage  := GetArgumentString(Arguments,'GENECHAINAGE');
    Generation_Affaire  := GetArgumentString(Arguments,'GENEAFFAIRE');
    Generation_Ressource  := GetArgumentString(Arguments,'GENERESSOURCE');
    if Generation_VenAch = 'ACH' then
       stProduitpgi:='GRF' else stProduitpgi:='GRC';
    end
  else
    begin
    Repeat
        Critere:=uppercase(ReadTokenSt(Arguments)) ;
        if Critere='LISTEMENU' then
           begin
           SetControlVisible('BATTACHACTION' ,True) ;
           SetControlVisible('BNEWACTION' ,True) ;
           end;
        if Critere='FICHETIERS' then
           TMenuItem(GetControl('mnTiers')).visible:=false;
        if Critere<>'' then
            begin
            x:=pos('=',Critere);
            if x<>0 then
               begin
               ChampMul:=copy(Critere,1,x-1);
               ValMul:=copy(Critere,x+1,length(Critere));
               if ChampMul='RCH_TIERS' then stTiers := ValMul ;
               if (ChampMul='TERMINE') and (ValMul='-') then SetControlEnabled ('RCH_TERMINE',False);
               if ChampMul='DERDATE' then stDerDate := ValMul ;
               if ChampMul='PRODUITPGI' then
                  stProduitpgi := ValMul;
               if ChampMul='TOBACTIONS' then
                  TobActions := TOB(StrToInt(ValMul));
               if ChampMul='FROMPIECE' then
                  TobPiece := TOB(StrToInt(ValMul));
               end;
            end;
    until  Critere='';

      Termine:=TCheckBox(GetControl('RCH_TERMINE')) ;
      if Assigned(Termine) then
        Termine.OnClick:=OnClickTermine;

      LesColonnesPiece :='FIXED;GP_NATUREPIECEG;GP_NUMERO;GP_DATEPIECE;GP_TOTALHT;GP_DEVISE;GP_SOUCHE;GP_INDICEG;GP_NATUREPIECEG;GP_ETATVISA;GP_NATUREPIECEG';
      GP:=THGRID(GetControl('GPIECES'));
      if Assigned(GP) then
        begin
        //GP.OnRowEnter:=GPRowEnter ;
        //GP.OnRowExit:=GPRowExit ;
        GP.OnDblClick:=GPLigneDClick ;
        GP.ColCount:=11;
        GP.PostDrawCell:= DessineCellGP;
        GP.ColWidths[0]:=4;
        GP.ColWidths[1]:=50; GP.ColFormats[1]:='CB=GCNATUREPIECEG';
        GP.ColWidths[2]:=50;
        GP.ColWidths[3]:=50; GP.ColAligns[3]:=tacenter;
        GP.ColWidths[4]:=50; GP.ColFormats[4]:='###0.00'; GP.ColAligns[4]:=taRightJustify ;
        GP.ColWidths[5]:=50; GP.ColFormats[5]:='CB=TTDEVISE';

        GP.ColWidths[6]:=0;
        GP.ColWidths[7]:=0;
        GP.ColWidths[8]:=0;
        GP.ColWidths[9]:=0;
        GP.ColWidths[10]:=0;
        GP.options:=GP.Options-[goEditing] ;
        TFFiche(Ecran).Hmtrad.ResizeGridColumns(GP) ;
      end;

      if Ecran <> nil then
        TFFiche(Ecran).OnKeyDown:=FormKeyDown ;

    if stProduitpgi = 'GRF' then
      begin
      SetControlText ('NATUREAUXI','FOU');
      SetControlVisible ('RCH_TABLELIBRECH1',false);
      SetControlVisible ('TRCH_TABLELIBRECH1',false);
      SetControlVisible ('RCH_TABLELIBRECH2',false);
      SetControlVisible ('TRCH_TABLELIBRECH2',false);
      SetControlVisible ('RCH_TABLELIBRECH3',false);
      SetControlVisible ('TRCH_TABLELIBRECH3',false);
      SetControlProperty ('RCH_CHAINAGE', 'Plus', 'RPG_PRODUITPGI="GRF"');
      end
    else
      begin
      soRtchrappel:=GetParamsocSecur('SO_RTCHRAPPEL',True);
      SetControlVisible ('RCH_TABLELIBRECHF1',false);
      SetControlVisible ('TRCH_TABLELIBRECHF1',false);
      SetControlVisible ('RCH_TABLELIBRECHF2',false);
      SetControlVisible ('TRCH_TABLELIBRECHF2',false);
      SetControlVisible ('RCH_TABLELIBRECHF3',false);
      SetControlVisible ('TRCH_TABLELIBRECHF3',false);
      SetControlProperty ('RCH_CHAINAGE', 'Plus', 'RPG_PRODUITPGI="GRC"');
      end;
    if soRtchrappel = FALSE then
       begin
       SetControlEnabled('RCH_GESTRAPPEL',FALSE) ;
       SetControlEnabled('RCH_DELAIRAPPEL',FALSE) ;
       SetControlEnabled('TRCH_DELAIRAPPEL',FALSE) ;
       end
    else
       begin
       Tcb:=TCheckBox(GetControl('RCH_GESTRAPPEL'));
       if Assigned(Tcb) then
         Tcb.OnClick:=RappelClick;
       end;
  {$IFDEF AFFAIRE}
     bAff:=TToolBarButton97(GetControl('BSELECTAFF1'));
     if Assigned(bAff) then
       bAff.OnClick:=RechercheAffaire;
     bAff:=TToolBarButton97(GetControl('BEFFACEAFF1'));
     if Assigned(bAff) then
       bAff.OnClick:=EffaceAffaire;

     if ( not (ctxAffaire in V_PGI.PGIContexte) ) and ( not ( ctxGCAFF in V_PGI.PGIContexte) ) or
        ( stProduitpgi='GRF') then
  {$ENDIF}
        begin
        SetControlVisible ('BEFFACEAFF1',false); SetControlVisible ('BSELECTAFF1',false);
        SetControlVisible ('TRCH_AFFAIRE0',false); SetControlVisible ('RCH_AFFAIRE1',false);
        SetControlVisible ('RCH_AFFAIRE2',false); SetControlVisible ('RCH_AFFAIRE3',false);
        SetControlVisible ('RCH_AVENANT',false);
        end;
    end;

    if stProduitpgi = 'GRF' then
      begin
      soRtchrappel:=GetParamsocSecur('SO_RFCHRAPPEL',True);
      soRtmesslibact:=GetParamSocSecur ('SO_RFMESSLIBACT',True);
      soRtmessbl:=GetParamSocSecur ('SO_RFMESSBL',True);
      soRtchchoixrappel:=GetParamsocSecur('SO_RFCHCHOIXRAPPEL','024');
      soRtgestinfos001:=false;
      soRtmesstypeact:=GetParamSocSecur ('SO_RFMESSTYPEACT',True);
      end
    else
      begin
      soRtchrappel:=GetParamsocSecur('SO_RTCHRAPPEL',True);
      soRtmesslibact:=GetParamSocSecur ('SO_RTMESSLIBACT',True);
      soRtmessbl:=GetParamSocSecur ('SO_RTMESSBL',True);
      soRtchchoixrappel:=GetParamsocSecur('SO_RTCHCHOIXRAPPEL','024');
      soRtgestinfos001:=GetParamSocSecur('SO_RTGESTINFOS001',False);
      soRtmesstypeact:=GetParamSocSecur ('SO_RTMESSTYPEACT',True);
      end;
    if TobPiece <> nil then SetControlVisible ('BACTIONS',true);

{$Ifdef GIGI}
 If (Not VH_GC.GaSeria) or not (GetParamSocSecur ('SO_AFRTPROPOS',False)) then
   begin
   if (GetControl('TRCH_AFFAIRE0') <> nil) then  SetControlVisible('TRCH_AFFAIRE0',false);
   if (GetControl('RCH_AFFAIRE1') <> nil) then  SetControlVisible('RCH_AFFAIRE1',false);
   if (GetControl('RCH_AFFAIRE2') <> nil) then  SetControlVisible('RCH_AFFAIRE2',false);
   if (GetControl('RCH_AFFAIRE3') <> nil) then  SetControlVisible('RCH_AFFAIRE3',false);
   if (GetControl('RCH_AVENANT') <> nil) then  SetControlVisible('RCH_AVENANT',false);
   if (GetControl('BEFFACEAFF1') <> nil) then  SetControlVisible('BEFFACEAFF1',false);
   if (GetControl('BSELECTAFF1') <> nil) then  SetControlVisible('BSELECTAFF1',false);
   end;
 SetControlVisible ('PIECES',False);// on en voit pas ongelt pièce
 SetControlvisible ('BPIECE',False);
{$endif}
if (GetParamSocSecur('SO_RTGESTIONGED',False) = False) or (stProduitpgi = 'GRF') then SetControlvisible ('BGED',False);
if Assigned(GetControl('BGED')) then
   TToolbarButton97(GetControl('BGED')).OnClick := RTAppelGED_OnClick;
if Assigned(GetControl('BDOCGEDEXIST')) then
   TToolbarButton97(GetControl('BDOCGEDEXIST')).OnClick := RTAppelGED_OnClick;
if OrigineIntervention then
  begin
  SetControlvisible ('BINSERT',False);
  SetControlvisible ('BDELETE',False);
  end;
{$IFNDEF EAGLSERVER}
  if Assigned(GetControl('MnAlerte')) then
    if AlerteActive('RCH') then
      TMenuItem(GetControl('MnAlerte')).OnClick := Alerte_OnClick_RCH
    else
      TMenuItem(GetControl('MnAlerte')).visible:=false;

  if Assigned(GetControl('MnListAlerte')) then
    if AlerteActive('RCH') then
      TMenuItem(GetControl('MnListAlerte')).OnClick := ListAlerte_OnClick_RCH
    else
      TMenuItem(GetControl('MnListAlerte')).visible:=false;

  if Assigned(GetControl('MnGestAlerte')) and Assigned(GetControl('MnAlerte'))
     and Assigned(GetControl('MnListAlerte')) then
         TMenuItem(GetControl('MnGestAlerte')).visible := (TMenuItem(GetControl('MnAlerte')).visible)
          and (TMenuItem(GetControl('MnListAlerte')).visible);
{$ENDIF EAGLSERVER}

end;

procedure TOM_ACTIONSCHAINEES.OnDeleteRecord;
var Requete : String;
begin
  inherited;
  Requete:='SELECT RAC_NUMACTION FROM ACTIONS WHERE RAC_AUXILIAIRE="'+TiersAuxiliaire(GetField('RCH_TIERS'), False)
               +'" AND RAC_NUMCHAINAGE='+IntToStr(GetField('RCH_NUMERO'));
  if ExisteSQL(Requete) then
  begin
    LastError:=2;
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
    exit ;
  end ;

  { GC/GRC : MNG / gestion des alertes }
{$IFNDEF EAGLSERVER}
  if (not V_Pgi.SilentMode) and (not AfterInserting)
    and (AlerteActive (TableToPrefixe(TableName))) then
      if (not ExecuteAlerteDelete(TForm(ecran),true)) then
        begin
        LastError := 99;
        exit;
        end;
{$ENDIF EAGLSERVER}

  ExecuteSQl('DELETE FROM CHAINAGEPIECES WHERE RLC_NUMCHAINAGE='+IntToStr(GetField('RCH_NUMERO'))+
      ' AND RLC_PRODUITPGI="'+stProduitpgi+'"' );
  if GetParamSocSecur('SO_RTGESTIONGED',False) then
     begin
     if ExisteSQL('SELECT RTD_DOCID FROM RTDOCUMENT WHERE RTD_TIERS="'+GetField('RCH_TIERS')+'" AND RTD_NUMCHAINAGE= '+IntToStr(GetField ('RCH_NUMERO'))) then
        ExecuteSQl('UPDATE RTDOCUMENT SET RTD_NUMCHAINAGE = 0 WHERE RTD_TIERS= "'+GetField('RCH_TIERS')
              +'" AND RTD_NUMCHAINAGE= '+ IntToStr(GetField('RCH_NUMERO')));
     end;
  if (Ecran<>nil) and ( ECran is TFFiche ) then
    TFFiche(Ecran).retour := IntToStr(GetField('RCH_NUMERO'))+';'+GetField('RCH_CHAINAGE');
end;

procedure TOM_ACTIONSCHAINEES.OnAfterUpdateRecord ;
begin
  Inherited ;
  // création des actions définies dans le type de chainage
  if Creation then
    CreationActionsChainages;
  if (Ecran<>nil) and ( Ecran is TFFiche ) and (NumChainage<>0) then
    TFFiche(Ecran).retour := IntToStr(NumChainage)+';'+
            iif(NumChainage=0,'',String(GetField('RCH_CHAINAGE')));
end ;


procedure TOM_ACTIONSCHAINEES.OnLoadRecord;
{$IFDEF AFFAIRE}
var    Q : TQuery;
{$ENDIF AFFAIRE}
begin
  inherited;
  {if not AfterInserting then
    LoadBufferAvantModif
  else
    VideBufferAvantModif;
  }
  If Not(DS.State in [dsInsert]) then
     begin
     ChargeGridPieces;
     SetControlEnabled ('RCH_CHAINAGE',False);
     end
  else
     SetControlEnabled ('RCH_CHAINAGE',true);
  stTermine:=GetField('RCH_TERMINE');
{$IFDEF AFFAIRE}
  if ( (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) ) and
       ( stProduitpgi <> 'GRF' ) then
    begin
    ChargeCleAffaire(THEDIT(GetControl('RCH_AFFAIRE0')),THEDIT(GetControl('RCH_AFFAIRE1')), THEDIT(GetControl('RCH_AFFAIRE2')),
      THEDIT(GetControl('RCH_AFFAIRE3')), THEDIT(GetControl('RCH_AVENANT')),Nil , taModif,GetField ('RCH_AFFAIRE'),True);
    if (ds.state in [dsinsert]) and (GetField ('RCH_AFFAIRE') <> '') then
      begin
      SetField('RCH_AFFAIRE0',THEDIT(GetControl('RCH_AFFAIRE0')).Text);
      SetField('RCH_AFFAIRE1',THEDIT(GetControl('RCH_AFFAIRE1')).Text);
      SetField('RCH_AFFAIRE2',THEDIT(GetControl('RCH_AFFAIRE2')).Text);
      SetField('RCH_AFFAIRE3',THEDIT(GetControl('RCH_AFFAIRE3')).Text);
      if THEDIT(GetControl('RCH_AVENANT')).Text ='' then THEDIT(GetControl('RCH_AVENANT')).Text:='00'; //mcd 22/12/2004
      SetField('RCH_AVENANT',THEDIT(GetControl('RCH_AVENANT')).Text);
      end;
    if (GetField ('RCH_AFFAIRE') <> '') then
      begin
      Q:= OpenSQL ('SELECT AFF_LIBELLE FROM AFFAIRE WHERE AFF_AFFAIRE="'+Getfield('RCH_AFFAIRE')+'"',True);
      if (Q<>Nil) and (not Q.Eof) then
        SetControlText ('TRCH_AFFAIRE',Q.FindField('AFF_LIBELLE').asstring);
      Ferme (Q);
      end
    else
      begin
      THEDIT(GetControl('RCH_AVENANT')).Text:='';
      SetField('RCH_AVENANT','');
      end;
    end;
{$ENDIF AFFAIRE}
  GestionBoutonGED;
end;

procedure TOM_ACTIONSCHAINEES.OnLoadAlerte;
begin
  {$IFNDEF EAGLSERVER}
  { GC/GRC : MNG / gestion des alertes }
  if (not V_Pgi.SilentMode) and (not AfterInserting)
    and (AlerteActive (TableToPrefixe(TableName))) then
      ExecuteAlerteLoad(TForm(Ecran),true);
  { GC/GRC : MNG / gestion des alertes }
  {$ENDIF !EAGLSERVER}
end;

{$IFDEF AFFAIRE}
procedure TOM_ACTIONSCHAINEES.RechercheAffaire (Sender: TObject);
var    Q : TQuery;
     stTiers : string;
    bProposition,NoChangeStatut : boolean;
begin
    { sauvegarde du code tiers qui disparait si aucune affaire n'ai sélectionnée }
    stTiers:=GetField('RCH_TIERS');
    bProposition:=false;
    NoChangeStatut:=VH_GC.GASeria;
    if stNatureauxi <> '' then
       begin
       if pos (stNatureauxi,FabriqueWhereNatureAuxiAff('PRO')) = 0 then
          begin
          bProposition:=false;
          NoChangeStatut:=false;
          end;
       if pos (stNatureauxi,FabriqueWhereNatureAuxiAff('AFF')) = 0 then
          begin
          bProposition:=true;
          NoChangeStatut:=false;
          end;
       end;
    GetAffaireEntete(THEDIT(GetControl('RCH_AFFAIRE')), THEDIT(GetControl('RCH_AFFAIRE1')),
                   THEDIT(GetControl('RCH_AFFAIRE2')), THEDIT(GetControl('RCH_AFFAIRE3')),
                   THEDIT(GetControl('RCH_AVENANT')),
                   THEDIT(GetControl('RCH_TIERS')), bProposition, NoChangeStatut, false, false, true,'');
    if (THEDIT(GetControl('RCH_AFFAIRE')).text <> '') then
    begin
    ForceUpdate;
    SetField('RCH_AUXILIAIRE',TiersAuxiliaire (THEDIT(GetControl('RCH_TIERS')).text, False));
    SetField('RCH_TIERS',THEDIT(GetControl('RCH_TIERS')).text); //mcd 22/12/2004 pour passer dans onchange
    SetField('RCH_AFFAIRE',THEDIT(GetControl('RCH_AFFAIRE')).Text);
    ChargeCleAffaire(THEDIT(GetControl('RCH_AFFAIRE0')),THEDIT(GetControl('RCH_AFFAIRE1')), THEDIT(GetControl('RCH_AFFAIRE2')),
       THEDIT(GetControl('RCH_AFFAIRE3')), THEDIT(GetControl('RCH_AVENANT')),Nil , taModif,THEDIT(GetControl('RCH_AFFAIRE')).text,True);

    SetField('RCH_AFFAIRE0',THEDIT(GetControl('RCH_AFFAIRE0')).Text);
    SetField('RCH_AFFAIRE1',THEDIT(GetControl('RCH_AFFAIRE1')).Text);
    SetField('RCH_AFFAIRE2',THEDIT(GetControl('RCH_AFFAIRE2')).Text);
    if  THEDIT(GetControl('RCH_AFFAIRE3')).Text ='' then THEDIT(GetControl('RCH_AFFAIRE3')).Text:='00'; //mcd 22/12/2004
    SetField('RCH_AFFAIRE3',THEDIT(GetControl('RCH_AFFAIRE3')).Text);
    if THEDIT(GetControl('RCH_AVENANT')).Text ='' then THEDIT(GetControl('RCH_AVENANT')).Text:='00'; //mcd 22/12/2004 si pas géré pas OK
    SetField('RCH_AVENANT',THEDIT(GetControl('RCH_AVENANT')).Text);
     Q:= OpenSQL ('SELECT AFF_LIBELLE FROM AFFAIRE WHERE AFF_AFFAIRE="'+THEDIT(GetControl('RCH_AFFAIRE')).text+'"',True);
     if (Q<>Nil) and (not Q.Eof) then
        SetControlText ('TRCH_AFFAIRE',Q.FindField('AFF_LIBELLE').asstring)
     else
        begin
        if Assigned(ecran) then SetFocusControl('BSELECTAFF1');
        LastError := 6;
        LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
        EffaceAffaire(Sender);
        end;
     Ferme(Q);
    end
    else
        begin
        ForceUpdate;
        SetField('RCH_TIERS',stTiers);
        end;
end;

procedure TOM_ACTIONSCHAINEES.EffaceAffaire (Sender: TObject);
begin
ForceUpdate;
SetControlText ('RCH_AFFAIRE',''); SetControlText ('RCH_AFFAIRE0',''); SetControlText('RCH_AFFAIRE1','');
SetControlText ('RCH_AFFAIRE2',''); SetControlText ('RCH_AFFAIRE3',''); SetControlText ('RCH_AVENANT','');
SetField ('RCH_AFFAIRE',''); SetField ('RCH_AFFAIRE0',''); SetField('RCH_AFFAIRE1','');
SetField ('RCH_AFFAIRE2',''); SetField ('RCH_AFFAIRE3',''); SetField ('RCH_AVENANT','');
SetControlText ('TRCH_AFFAIRE','');
end;

Procedure TOM_ACTIONSCHAINEES.AffaireEnabled( grise : boolean);
begin
  SetControlEnabled ('BEFFACEAFF1',grise); SetControlEnabled ('BSELECTAFF1',grise);
  SetControlEnabled ('RCH_AFFAIRE1',grise);
  SetControlEnabled ('RCH_AFFAIRE2',grise); SetControlEnabled ('RCH_AFFAIRE3',grise);
  SetControlEnabled ('RCH_AVENANT',grise);
end;
{$ENDIF}

procedure TOM_ACTIONSCHAINEES.OnNewRecord;
var TobTypeEncours: tob;
    Aff0, Aff1, Aff2, Aff3, Avenant: string;
begin
  inherited;
  VH_RT.TobTypesChainage.Load;

  if Generation_Tiers <> '' then
    begin
    stTiers:=Generation_Tiers;
    SetField('RCH_CHAINAGE',Generation_Chainage);
    SetField('RCH_AFFAIRE',Generation_Affaire);
    //AB-200706-FQ GA 14163
    CodeAffaireDecoupe (Generation_Affaire, Aff0, Aff1, Aff2, Aff3, Avenant, taCreat, false);
    SetField ('RCH_AFFAIRE0', Aff0);
    SetField ('RCH_AFFAIRE1', Aff1);
    SetField ('RCH_AFFAIRE2', Aff2);
    SetField ('RCH_AFFAIRE3', Aff3);
    SetField ('RCH_AVENANT', Avenant);
    TobTypeEncours:=VH_RT.TobTypesChainage.FindFirst(['RPG_CHAINAGE','RPG_PRODUITPGI'],[Generation_Chainage,stProduitpgi],TRUE) ;
    if TobTypeEncours <> Nil then
       SetField ('RCH_LIBELLE',TobTypeEncours.Getvalue('RPG_LIBELLE'))
    else
       SetField ('RCH_LIBELLE',Generation_LibTiers);
    end
  else
    begin
    if StTiers <> '' then
       SetControlEnabled ('RCH_TIERS',False)
    else
       SetControlEnabled ('RCH_TIERS',True);
    SetControlEnabled ('RCH_NUMERO',False);
    if Assigned(ecran) then SetFocusControl('RCH_LIBELLE');
    end;

  SetField('RCH_DATEDEBUT', V_PGI.DateEntree);
  SetField('RCH_DATEFIN', iDate2099);

  if (trim(stTiers)<>'') then
     begin
     setfield ('RCH_TIERS',stTiers);
     SetField('RCH_AUXILIAIRE',TiersAuxiliaire (GetField('RCH_TIERS'), False));
     end;

  if Generation_Ressource <> '' then
    SetField('RCH_INTERVENANT',Generation_Ressource)
  else
    SetField('RCH_INTERVENANT',VH_RT.RTResponsable);

  // ? mng à tester SetField('TRCH_INTERVENANT_',VH_RT.RTNomResponsable);
  if stProduitpgi='GRF' then
     SetField('RCH_PRODUITPGI','GRF')
  else
     SetField('RCH_PRODUITPGI','GRC');
end;

procedure TOM_ACTIONSCHAINEES.OnUpdateRecord;
var Q : TQuery;
    nb_chaine,j : Integer;
    HeureAct : TDateTime;
    TobActionsChainage : Tob;
    StPiecesVisees : String;
begin
  inherited;

  VH_RT.TobTypesChainage.Load;

  If (trim(GetField('RCH_LIBELLE'))='') then
  begin
    if Assigned(ecran) then SetFocusControl('RCH_LIBELLE');
    LastError:=1;
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
    exit ;
  end;
  TobActionsChainage:=Nil;
  If (GetField('RCH_CHAINAGE')<>'') then
  begin
    TobActionsChainage:=VH_RT.TobTypesChainage.FindFirst(['RPG_CHAINAGE','RPG_PRODUITPGI'],[GetField('RCH_CHAINAGE'),stProduitpgi],TRUE) ;
    if TobActionsChainage = Nil then
    begin
      if Assigned(ecran) then SetFocusControl('RCH_CHAINAGE');
      LastError:=4;
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      exit ;
    end;
  end;

  if (GetField('RCH_TIERS')='') then
    begin
    if Assigned(ecran) then SetFocusControl('RCH_TIERS');
    LastError := 3;
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
    exit;
    end;
  if GetField('RCH_AUXILIAIRE') = '' then
    begin
    if Assigned(ecran) then SetFocusControl('RCH_TIERS');
    LastError := 3;
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
    exit;
    end;
  if GetField ('RCH_DATEFIN') < GetField ('RCH_DATEDEBUT') then
     begin
     Lasterror:=8;
     LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
     SetFocusControl('RCH_DATEFIN') ;
     exit;
     end;

  { GC/GRC : MNG / gestion des alertes }
{$IFNDEF EAGLSERVER}
  if (ds<>nil) and (not V_Pgi.SilentMode) and
     (AlerteActive (TableToPrefixe(TableName))) then
      if (not ExecuteAlerteUpdate(TForm(Ecran),true)) then
        LastError := 99;
{$ENDIF EAGLSERVER}
  { GC/GRC : MNG fin / gestion des alertes }


  {if not ExisteSQL ('SELECT T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="'+GetField('RCH_AUXILIAIRE')+'"') then
  begin
  if Assigned(ecran) then SetFocusControl('RCH_TIERS');
  LastError := 3;
  LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
  exit;
  end;}

  if  (GetField ('RCH_GESTRAPPEL') = 'X' ) and (GetField ('RCH_DELAIRAPPEL')<>'') then
    begin
    HeureAct:=GetField ('RCH_DATEDEBUT');
    SetField('RCH_DATERAPPEL',PlusDate(HeureAct, (GetField ('RCH_DELAIRAPPEL')/24)* (-1),'J'));
    end;
  Creation:=False;
  if Generation_Tiers <> '' then
     Creation:=True
  else
    if Assigned(DS) and (DS.State = dsInsert)  then
      Creation:=True;
  if Creation then
    begin
    Q := OpenSQL('SELECT MAX(RCH_NUMERO) FROM ACTIONSCHAINEES WHERE RCH_PRODUITPGI="'+stProduitpgi+'"', True);
    if not Q.Eof then
      begin
      nb_chaine := Q.Fields[0].AsInteger;
      setfield ('RCH_NUMERO',nb_chaine + 1);
      end else
      setfield ('RCH_NUMERO',1);
    Ferme(Q) ;
    end;

  NumChainage:=GetField('RCH_NUMERO');
  if ({isFieldModified( 'RCH_TERMINE' )} stTermine<>'X') and ( Getfield('RCH_TERMINE' )='X' )
    and (GetField('RCH_CHAINAGE')<>'') and ( TobActionsChainage <> Nil ) then
  begin
    if Assigned(GP) and ( TobActionsChainage.GetValue('RPG_VISA') = 'X' ) then
    begin
      StPiecesVisees:='';
      for j:=1 to (GP.RowCount-1) do
      begin
        if (GP.Cells[9,j] = 'ATT' ) then
        begin
        SetPieceVisee (GP.Cells[10,j],GP.Cells[6,j],IntToStr(Valeuri(GP.Cells[2,j])),GP.Cells[7,j]);
          if StPiecesVisees = '' then
             StPiecesVisees:=TraduireMemoire(TexteMessage[7]);
          StPiecesVisees:=StPiecesVisees+#13#10+'  '+GP.Cells[1,j]+Traduirememoire(' n° : ')+GP.Cells[2,j]
        end;
      end;
      if StPiecesVisees<>'' then StPiecesVisees:=StPiecesVisees+#13#10;
    end;
    if ( TobActionsChainage.GetValue('RPG_MAILTERMINE') = 'X' ) and
       ( Getfield('RCH_INTERVENANT') <> '') and
       ( Getfield('RCH_INTERVENANT') <> VH_RT.RTResponsable ) then
      EnvoiMessage(Getfield('RCH_INTERVENANT')+'|'+Getfield('RCH_LIBELLE'),'-',TobActionsChainage.GetValue('RPG_MAILTERMAUTO'),true,Nil,StPiecesVisees);
    if StPiecesVisees <> '' then
       PgiInfo (StPiecesVisees,Ecran.Caption);
  end;
end;

procedure TOM_ACTIONSCHAINEES.OnChangeField(F: TField);
var TobTypeEncours : tob;
    Q : TQuery;
    betat : boolean;
    i : integer;
begin
inherited;
//  if (DS.State = dsInsert) and(F.FieldName = 'RCH_TIERS') and (GetField('RCH_TIERS')<>'') then
//      SetField('RCH_AUXILIAIRE',TiersAuxiliaire (GetField('RCH_TIERS'), False));
   if (F.FieldName = 'RCH_TIERS') and (GetField('RCH_TIERS')<>'') then
      begin
      Q := OpenSQL('SELECT T_NATUREAUXI,T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+GetField('RCH_TIERS')+'"', True);
      if Q.Eof then
        begin
        Ferme(Q);
        if Assigned(ecran) then SetFocusControl('RCH_TIERS');
        LastError := 3;
        LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
        exit;
        end
      else
        begin
        stNatureauxi:=Q.FindField('T_NATUREAUXI').asstring;
        if (DS.State = dsInsert) then
           SetField('RCH_AUXILIAIRE',Q.FindField('T_AUXILIAIRE').asstring);
{$IFDEF AFFAIRE}
        if ( (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) ) then
          begin
          if ( pos (stNatureauxi, FabriqueWhereNatureAuxiAff('AFF')) = 0 ) and
             ( ( pos (stNatureauxi, FabriqueWhereNatureAuxiAff('PRO')) = 0 ) or
             (VH_GC.GASeria=false) ) then
             AffaireEnabled(false)
          else
             AffaireEnabled(true);
          end;
{$ENDIF}
        end;
      Ferme(Q);
      end;

  if (F.FieldName = 'RCH_CHAINAGE') and (GetField('RCH_CHAINAGE')<>'') then
  begin
    VH_RT.TobTypesChainage.Load;

    TobTypeEncours:=VH_RT.TobTypesChainage.FindFirst(['RPG_CHAINAGE','RPG_PRODUITPGI'],[GetField('RCH_CHAINAGE'),stProduitpgi],TRUE) ;
    if TobTypeEncours <> Nil then
    begin
      if (DS.State = dsInsert) then
        begin
        if stProduitpgi='GRC' then
          begin
          Setfield ('RCH_TABLELIBRECH1',TobTypeEncours.Getvalue('RPG_TABLELIBRECH1'));
          Setfield ('RCH_TABLELIBRECH2',TobTypeEncours.Getvalue('RPG_TABLELIBRECH2'));
          Setfield ('RCH_TABLELIBRECH3',TobTypeEncours.Getvalue('RPG_TABLELIBRECH3'));
          end
        else
          begin
          Setfield ('RCH_TABLELIBRECHF1',TobTypeEncours.Getvalue('RPG_TABLELIBRECHF1'));
          Setfield ('RCH_TABLELIBRECHF2',TobTypeEncours.Getvalue('RPG_TABLELIBRECHF2'));
          Setfield ('RCH_TABLELIBRECHF3',TobTypeEncours.Getvalue('RPG_TABLELIBRECHF3'));
          end;
        if GetField('RCH_LIBELLE')='' then
           SetField ('RCH_LIBELLE',TobTypeEncours.Getvalue('RPG_LIBELLE'));
        end;
      if TobTypeEncours.Getvalue('RPG_NONTERMINE')='X' then
         begin
         betat:=true;
         if (Assigned(TobActions)) and (TobActions.detail.count <> 0 ) then
            for i := 0 to Pred(TobActions.detail.count) Do
             if (TobActions.detail[i].GetValue('RAC_ETATACTION') <> REALISEE ) and
                (TobActions.detail[i].GetValue('RAC_ETATACTION') <> ANNULEE ) then
                begin
                betat:=false;
                break;
                end;
         SetControlEnabled ('RCH_TERMINE',betat);
         end;
    end
    else
    begin
      if Assigned(ecran) then SetFocusControl('RCH_CHAINAGE');
      LastError:=4;
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      exit ;
    end;
  end;
  if GetCheckBoxState('RCH_TERMINE') = cbChecked then
      begin
      SetControlEnabled ('BPIECE',False);
      SetControlEnabled ('BATTACHACTION',False);
      SetControlEnabled ('BNEWACTION',False);
      end
   else
      begin
      SetControlEnabled ('BPIECE',True);
      SetControlEnabled ('BATTACHACTION',True);
      SetControlEnabled ('BNEWACTION',True);
      end;
end;

procedure TOM_ACTIONSCHAINEES.OnClose;
begin
inherited;
if (Ecran<>nil) and ( ECran is TFFiche ) and (GetField('RCH_NUMERO') <> Null )
   and (TFFiche(Ecran).retour = '' ) then
       TFFiche(Ecran).retour := IntToStr(NumChainage)+';'+
            iif(NumChainage=0,'',String(GetField('RCH_CHAINAGE')));
end;

procedure TOM_ACTIONSCHAINEES.OnClickTermine(Sender: TObject);
var DateAct : TDateTime;
begin
if (stDerDate<>DateToStr(iDate1900)) and (GetCheckBoxState('RCH_TERMINE') = cbChecked) then
    begin
    DateAct:=StrToDate(stDerDate);
    SetField ('RCH_DATEFIN', DateAct);
    end;
if GetCheckBoxState('RCH_TERMINE') = cbChecked then
  begin
  SetControlEnabled ('BPIECE',False);
  SetControlEnabled ('BATTACHACTION',False);
  SetControlEnabled ('BNEWACTION',False);
//  TFFiche(Ecran).retour := 'X';  
  end
else
  begin
  SetControlEnabled ('BPIECE',True);
  SetControlEnabled ('BATTACHACTION',True);
  SetControlEnabled ('BNEWACTION',True);
  end

end;

procedure DessineCell (GG : Thgrid; ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
  Arect: Trect ;
begin
If Arow < GG.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=GG.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := GG.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = GG.row) then
         BEGIN
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=ARect.Right-2 ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         END ;
    end;
end;

procedure TOM_ACTIONSCHAINEES.DessineCellGP ( ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
begin
DessineCell( GP,ACol,ARow,Canvas,AState);
end;

procedure TOM_ACTIONSCHAINEES.GPLigneDClick (Sender: TObject);
var
  CleDoc: R_CleDoc;
  SaveAutoSave: boolean;
  Param: R_SaisiePieceParam;
begin
if GP.Cells[2, GP.Row] <> '' then
    begin
    // on valide d'abord sinon problème de multi à la prochaine validation

    if not TFFiche(Ecran).Bouge(TNavigateBtn(nbPost)) then exit ;
    DecodeRefPiece (
        GP.Cells[8, GP.Row] + ';' + GP.Cells[3, GP.Row] + ';' + GP.Cells[6,GP.Row] + ';' +
             IntToStr(Valeuri(GP.CellValues[2,GP.Row])) +
        ';' + GP.CellValues[7,GP.Row] + ';',
        CleDoc);
    { si l'on vient de la pièce, pas accès à cette pièce par double clic }
    if TobPiece <> nil then
      begin
      if ( TOBPiece.GetString('GP_NATUREPIECEG') = CleDoc.Naturepiece ) and
         ( TOBPiece.GetString('GP_SOUCHE') = CleDoc.Souche ) and
         ( TOBPiece.GetInteger('GP_NUMERO') = CleDoc.NumeroPiece ) and
         ( TOBPiece.GetInteger('GP_INDICEG') = CleDoc.Indice ) then
         exit; 
      end;
//    InitRecordSaisiePieceParam(Param);

    SaisiePiece(CleDoc, taConsult);
    //DS.Refresh ;
    SaveAutoSave:=TFFiche(ecran).FAutoSave.checked ;
    TFFiche(ecran).FAutoSave.checked:=True;
    TFFiche(ecran).OM.RefreshDB ;
    TFFiche(ecran).FAutoSave.checked:=SaveAutoSave;
{$IFDEF EAGLCLIENT}
    ChargeGridPieces;
{$ENDIF}
    end;
end;

procedure TOM_ACTIONSCHAINEES.RTChainageAttacheAction;
begin
if not TFFiche(Ecran).Bouge(TNavigateBtn(nbPost)) then exit ;
AglLancefiche('RT','RTACTIONS_TIERS','RAC_TIERS='+GetField('RCH_TIERS')+';RAC_NUMCHAINAGE=0;NUMCHAINAGE='+IntToStr(GetField('RCH_NUMERO')),'','NOCHANGEPROSPECT;NOCHANGECHAINAGE') ;
OnloadRecord;
end;

procedure TOM_ACTIONSCHAINEES.ChainageNouvellePiece;
var StArg : String;
    TobActionsChainage : tob;
begin
  StArg:='';
  { si envoi mail, on demande une confirmation }
  if DS.State= dsInsert then
  begin
    VH_RT.TobTypesAction.Load;

    TobActionsChainage:=VH_RT.TobTypesAction.FindFirst(['RPA_CHAINAGE','RPA_PRODUITPGI'],[GetField('RCH_CHAINAGE'),stProduitpgi],TRUE) ;
    if TobActionsChainage <> Nil then
       if TobActionsChainage.GetValue('RPA_MAILETATVAL') = 'X' then
          if PGIAsk(TraduireMemoire('Un message va être envoyé, êtes-vous sur d''avoir bien renseigné le bloc-notes?'),ecran.caption)=mrno then
             exit;
  end;

  {if DS.State= dsInsert then TFFiche(Ecran).Caption:= 'Chainage créé'
     else TFFiche(Ecran).Caption:= 'Chainage enregistré';}

  if not TFFiche(Ecran).Bouge(TNavigateBtn(nbPost)) then exit;

  if VH_GC.GCIfDefCEGID then
    StArg:='CEGID;';
  if OrigineIntervention then
    StArg:=StArg+'ORIGINEINTERVENTION;';
  if stProduitpgi = 'GRC' then
    begin
    if OrigineIntervention and not ExistPieceSAV then exit;
    AGLLanceFiche ('RT','RTACTCHA_MUL_PIEC','GP_TIERS='+getField('RCH_TIERS')
       +';NUMCHAINAGE='+intToStr(getField('RCH_NUMERO')),'',StArg);
    end
  else
    AGLLanceFiche ('RT','RFACTCHA_MUL_PIEC','GP_TIERS='+getField('RCH_TIERS')
       +';NUMCHAINAGE='+intToStr(getField('RCH_NUMERO')),'',StArg);

OnloadRecord;
end;

procedure TOM_ACTIONSCHAINEES.RTGenereLotActions;
var TobAction,TobTypeEncours,TobTypActEncours : TOB;
    ListeResp,MailAuto,StTypeAction,StArg : String;
    i :integer;
begin
  if getField('RCH_TIERS') = '' then
    begin
    if Assigned(ecran) then SetFocusControl('RCH_TIERS');
    LastError := 5;
    PGIInfo(TraduireMemoire(TexteMessage[LastError]),TraduireMemoire('Générer un lot d''actions')) ;
    exit;
    end;
  if not TFFiche(Ecran).Bouge(TNavigateBtn(nbPost)) then exit;
  TobAction:=TOB.create ('ACTIONS',NIL,-1);
  TobAction.PutValue ('RAC_BLOCNOTE',GetControlText('RCH_BLOCNOTE'));
  TobAction.PutValue ('RAC_AFFAIRE',GetControlText('RCH_AFFAIRE'));
  TobAction.PutValue ('RAC_AFFAIRE0',GetControlText('RCH_AFFAIRE0'));
  TobAction.PutValue ('RAC_AFFAIRE1',GetControlText('RCH_AFFAIRE1'));
  TobAction.PutValue ('RAC_AFFAIRE2',GetControlText('RCH_AFFAIRE2'));
  TobAction.PutValue ('RAC_AFFAIRE3',GetControlText('RCH_AFFAIRE3'));
  TobAction.PutValue ('RAC_AVENANT',GetControlText('RCH_AVENANT'));

  TheTob := TobAction;
  ListeResp:=AGLLanceFiche('RT','RTACTIONS_LOT','','','TIERSCHAINAGE='+getField('RCH_TIERS')+
           ';NUMCHAINAGE='+intToStr(getField('RCH_NUMERO'))+';PRODUITPGI='+stProduitpgi);
  { 3ème argument=Type action }
  StArg:=ListeResp;
  for i:=1 to 3 do
     StTypeAction :=ReadToKenPipe(StArg,'|');

  VH_RT.TobTypesAction.Load;

  TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[StTypeAction,'---',0],TRUE) ;
  if (TobTypActEncours <> Nil) and (not V_Pgi.SilentMode) then
     MailAuto:=TobTypActEncours.GetValue('RPA_MAILVALAUTO')
  else
     MailAuto:='-';

  FreeAndNil(TobAction);
  OnloadRecord;
  if Assigned(ecran) then SetFocusControl('ACTIONS');
  if ListeResp <> '' then
  begin
    VH_RT.TobTypesChainage.Load;

    TobTypeEncours:=VH_RT.TobTypesChainage.FindFirst(['RPG_CHAINAGE','RPG_PRODUITPGI'],[GetField('RCH_CHAINAGE'),stProduitpgi],TRUE) ;

    if (TobTypeEncours<>Nil ) then
      EnvoiMessage(ListeResp, TobTypeEncours.GetValue('RPG_MENTIONBLC'),MailAuto,false,Nil,'')
    else
      EnvoiMessage(ListeResp, '-',MailAuto,false,Nil,'');
  end;
end;

procedure TOM_ACTIONSCHAINEES.EnvoiMessage (ListeResp,MentionBlc,MailAuto:String; Termine : boolean; St : HTStringList; PiecesVisees : String);
var StDestinataire,StSQL,StWhere,Critere,Objet,StArg,StLibelle,StNumChainage,StRessources: string;
    i : integer;
    Q : TQuery;
    TobIntervint:TOB;
    Memo : TMemo;
    Liste :HTStringList;
begin
  StArg:=ListeResp;
  StDestinataire :=ReadToKenPipe(StArg,'|');
  StLibelle:=ReadToKenPipe(StArg,'|');
  StNumChainage:=ReadToKenPipe(StArg,'|');
  Stwhere := 'WHERE ARS_RESSOURCE in (';
  i:=0;
  While StDestinataire <> '' do
  begin
      inc(i);
      Critere :=ReadTokenSt(StDestinataire);
      if i = 1 then
        StWhere := StWhere + '"'+Critere+'"'
      else
        StWhere := StWhere + ',"'+Critere+'"'
  end;
  StSQL:='SELECT ARS_RESSOURCE,ARS_LIBELLE,ARS_LIBELLE2,ARS_TYPERESSOURCE,ARS_EMAIL'+
             ' FROM RESSOURCE ';
  if StWhere<>'' then StSQL:=StSQL+stWhere+')';

  TobIntervint:=TOB.create ('les intervenants',NIL,-1);
  Q := OpenSql (StSQL,TRUE);
  TobIntervint.LoadDetailDB ('RESSOURCE', '', '', Q, False);
  ferme(Q) ;

  StDestinataire := ''; StRessources:='';

  for i:=0 to TobIntervint.detail.count-1 do
  begin
    if TobIntervint.detail[i].GetValue('ARS_EMAIL') <> '' then
    begin
      StDestinataire := StDestinataire + TobIntervint.detail[i].GetValue('ARS_EMAIL') +';';
      { sur option du type chainage, on mentionne dans le Blc l'envoi d'email }
      if (MentionBlc= 'X' ) then
          StRessources:=StRessources+' '+TobIntervint.detail[i].GetValue('ARS_LIBELLE2')
             +' '+TobIntervint.detail[i].GetValue('ARS_LIBELLE')+',';
    end;
  end;

  if StDestinataire <> '' then
  begin
   Objet := '';
   if St=Nil then
      if GetControlText('TRCH_TIERS_') <> '' then
        Objet := 'Client : '+GetControlText('TRCH_TIERS_');

   if soRtmesslibact = True  then
      begin
      if not Termine then
         StLibelle:=' '+StLibelle+' pour '+GetField('RCH_LIBELLE')+' no '+IntToStr(GetField('RCH_NUMERO'))
      else
         if St=Nil then
            StLibelle:=' '+StLibelle+' no '+IntToStr(GetField('RCH_NUMERO'))
         else
            StLibelle:=' '+StLibelle+' no '+StNumChainage;

      if Objet <> '' then Objet := Objet + ' - ' + StLibelle
      else Objet := StLibelle;
      end;

   if soRtmessbl = True then
      begin
      if St=Nil then
        begin
        Memo := TMemo(GetControl('RCH_BLOCNOTE'));
        Liste:=HTStringList.Create ;
        Liste.Text := Memo.text;
//        nblignes:=Memo.Lines.Count;
        if Termine then
          begin
            if PiecesVisees = '' then
//              Memo.Lines.Insert(0,'Chainage Terminé :'+#13#10)
                Liste.Insert(0,'Chaînage Terminé :'+#13#10)
            else
              begin
//              Memo.Lines.Insert(0,PiecesVisees);
//              nblignes:=Memo.Lines.Count-nblignes;
              Liste.Insert(0,PiecesVisees);
              end;
          end;
//        SendMail (Objet,StDestinataire,'',TStringList(Memo.lines),'',(MailAuto='X'),1,'','');
        PGIEnvoiMail (Objet,StDestinataire,'',Liste,'',(MailAuto='X'),1,'','');
{        if Termine then
          begin
          if PiecesVisees <> ''
             then for i:=0 to (nblignes-1) do Memo.Lines.Delete(0)
             else for i:=0 to 1 do Memo.Lines.Delete(0);
          end;      }
        Liste.free;
        end
      else
        PGIEnvoiMail (Objet,StDestinataire,'',St,'',(MailAuto='X'),1,'','');
      end
   else
      PGIEnvoiMail (Objet,StDestinataire,'',Nil,'',(MailAuto='X'),1,'','');
    { alimentation du bloc-notes }
    if (MentionBlc= 'X' ) and (StRessources<>'') then
    begin
     DS.Edit;
     Memo := TMemo(GetControl('RCH_BLOCNOTE'));
     if copy(StRessources,length(StRessources),1) = ',' then delete(StRessources,length(StRessources),1);
     Memo.Lines.Add(TraduireMemoire('email envoyé le ')+DateTimeToStr(Now)+TraduireMemoire(' à')+
        StRessources);
     SetField ('RCH_BLOCNOTE',Memo.Text);
    end
  end
  else
   begin
   if (not V_Pgi.SilentMode) then
   PGIBox('Envoi de message impossible : pas d''adresse e-mail','Chaînage d''actions');
   end;
  FreeAndNil(TobIntervint);
end;

procedure TOM_ACTIONSCHAINEES.ChargeGridPieces ;
var QQ : TQuery;
Tobpiece : TOB;
Requete : String;
begin
TobPiece:=tob.create('_PIECE',Nil,-1) ;
Requete:='SELECT GP_NATUREPIECEG,GP_NUMERO,GP_DATEPIECE,GP_TOTALHT,GP_DEVISE,GP_SOUCHE,GP_INDICEG,GP_ETATVISA from CHAINAGEPIECES '+
'left join PIECE on GP_NUMERO=RLC_NUMERO and GP_SOUCHE=RLC_SOUCHE and GP_INDICEG=RLC_INDICEG '+
'and GP_NATUREPIECEG=RLC_NATUREPIECEG'+
' where RLC_NUMCHAINAGE='+IntToStr(getfield('RCH_NUMERO'))+' AND RLC_PRODUITPGI="'+stProduitpgi+'"';

QQ:=OpenSQL(Requete,True);
TobPiece.LoadDetailDB('PIECE','','',QQ,false,true) ;
ferme(QQ);
TobPiece.PutGridDetail(GP,True,True,LesColonnesPiece,True);
FreeAndNil(TobPiece) ;
end;

procedure TOM_ACTIONSCHAINEES.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  Case Key of
  VK_DELETE :
    BEGIN
    if (Screen.ActiveControl = GP) and (Shift=[ssCtrl]) and
        ( GP.Cells[1,GP.Row] <> '' ) then
       begin
       ExecuteSQl('delete from chainagepieces where rlc_naturepieceg="'+GP.Cells[8,GP.row]+'" and '+
       'rlc_souche="'+GP.Cells[6,GP.row]+'" and rlc_numero='+IntToStr(Valeuri(GP.Cells[2,GP.row]))+
       ' and rlc_indiceg='+GP.Cells[7,GP.row]+' and rlc_numchainage='+IntToStr(getfield('RCH_NUMERO'))+
       ' AND RLC_PRODUITPGI="'+stProduitpgi+'"');
       if GP.RowCount = 2 then
          GP.videpile (False)
       else
          GP.DeleteRow (GP.Row);
       end
    END;
    VK_F3 :{GED}
    BEGIN
    if (ssAlt in Shift) then
       if (GetParamSocSecur('SO_RTGESTIONGED',False) = True) and (stProduitpgi = 'GRC') then
         begin
            if (ds<>nil) and not(DS.State in [dsInsert]) then
              begin
              Key:=0 ;
              RTAppelGED_OnClick (Sender);
              end;
         end;
    END;
    81 : {Ctrl + Q - Création d'1 alerte} if (ssCtrl in Shift) then
          begin
          Key:=0 ;
          Alerte_OnClick_Rch(Sender);
          end;
    85 : {Ctrl + U - liste des alertes du tiers} if (ssCtrl in Shift) then
          begin
          Key:=0 ;
          ListAlerte_OnClick_Rch(Sender);
          end;
   end;
end;

procedure TOM_ACTIONSCHAINEES.RappelClick(Sender: TObject);
begin
if (GetCheckBoxState('RCH_GESTRAPPEL') = cbChecked) and (DS.State in [dsInsert,dsEdit]) then
    SetField('RCH_DELAIRAPPEL',soRtchchoixrappel)
else
    SetField('RCH_DELAIRAPPEL','');
end;

procedure TOM_ACTIONSCHAINEES.CreationActionsChainages;
var TobAction,{TobActionSav,}TobActionsChainage,TobTypeEncours,TobActionsLocal : Tob;
    nbact,i,NumSem, AnneeSem : integer;
    DateAction,vDtCurDate : TDateTime;
    Select,EnvoiMail,MailResp,MailAuto,OutLook : string;
    Q : TQuery;
    Duree : Word;
begin
  nbact:=0;

  Select := 'SELECT MAX(RAC_NUMACTION) FROM ACTIONS WHERE RAC_AUXILIAIRE = "'+ GetField('RCH_AUXILIAIRE')+'"';
  Q:=OpenSQL(Select, True);
  if not Q.Eof then
     nbact := Q.Fields[0].AsInteger;
  Ferme(Q) ;

  DateAction:=GetField ('RCH_DATEDEBUT');
  // chargement des actions du chainage à créer : elles ont été chargées ds l'ordre :
  // Code chainage / No ligne
  VH_RT.TobTypesAction.Load;
  TobActionsLocal := Tob.Create('prise locale', nil, -1); // VH_RT.TobTypesAction est utilisée ds RTCreationLiensActions
  TobActionsLocal.Dupliquer(VH_RT.TobTypesAction, true, True);

  TobActionsChainage:=TobActionsLocal.FindFirst(['RPA_CHAINAGE','RPA_PRODUITPGI'],[GetField('RCH_CHAINAGE'),stProduitpgi],false) ;
  if TobActionsChainage = Nil then exit;

  VH_RT.TobTypesChainage.Load;

  {TobActionSav :=TOB.create ('ACTIONS',NIL,-1);}
  i:=0;
  while Assigned(TobActionsChainage) do
    begin
    //if i = 0 then
       begin
       EnvoiMail:= TobActionsChainage.GetValue('RPA_MAILETATVAL');
       MailResp := TobActionsChainage.GetValue('RPA_MAILRESPVAL');
       MailAuto := TobActionsChainage.GetValue('RPA_MAILVALAUTO');
       if ( V_Pgi.SilentMode) then
          MailAuto:='-';
       OutLook := TobActionsChainage.GetValue('RPA_OUTLOOK');
       end;
    TobAction:=TOB.create ('ACTIONS',NIL,-1);
    TobAction.InitValeurs ;
    Inc(nbact);
    TobAction.PutValue ('RAC_NUMACTION',nbact);
    TobAction.PutValue ('RAC_AUXILIAIRE',GetField('RCH_AUXILIAIRE'));
    TobAction.PutValue ('RAC_TIERS',GetField('RCH_TIERS'));
    TobAction.PutValue ('RAC_NUMCHAINAGE',GetField('RCH_NUMERO'));
    TOBCopieChamp(TobActionsChainage,TobAction);
    TobAction.putvalue ('RAC_DATECREATION', Now);
    TobAction.putvalue ('RAC_DATEMODIF', Now);
    TobAction.putvalue ('RAC_CREATEUR', V_PGI.User);
    TobAction.putvalue ('RAC_UTILISATEUR', V_PGI.User);
    if i = 0 then TobAction.PutValue ('RAC_INTERVINT',TobAction.GetValue ('RAC_INTERVENANT')+';');
    if TobAction.GetValue ('RAC_DUREEACTION') = 0 then
      TobAction.PutValue ('RAC_DUREEACTION',30);

    DateAction:=PlusDate(DateAction, StrToInt(TobActionsChainage.GetValue('RPA_NBJOURS')) ,'J');
    { sur option du type chainage, on saute ou non les week-ends }
    TobTypeEncours:=VH_RT.TobTypesChainage.FindFirst(['RPG_CHAINAGE','RPG_PRODUITPGI'],[GetField('RCH_CHAINAGE'),stProduitpgi],TRUE) ;
    if (TobTypeEncours <> Nil) then
        if (TobTypeEncours.GetValue('RPG_WEEKEND')= 'X' ) then
    begin
      NumSem := NumSemaine(DateAction,AnneeSem);
      vDtCurDate := PremierJourSemaine(NumSem, AnneeSem);
      if DateAction=PlusDate(vDtCurDate,5,'J') then
          DateAction:=PlusDate(DateAction,2,'J')
      else
          if DateAction=PlusDate(vDtCurDate,6,'J') then
              DateAction:=PlusDate(DateAction,1,'J');
    end;
    TobAction.PutValue ('RAC_DATEACTION',DateAction);
    TobAction.PutValue('RAC_DATEECHEANCE',iDate2099);
    if TobActionsChainage.GetValue('RPA_DELAIDATECH') <> 0 then TobAction.PutValue('RAC_DATEECHEANCE',RTCalculEch(DateAction,StrToInt(TobActionsChainage.GetValue('RPA_DELAIDATECH')),TobActionsChainage.GetValue('RPA_WEEKEND')));

    if  (TobAction.GetValue('RAC_GESTRAPPEL') = 'X' ) then
        begin
        // calcul différent suivant que l'on traite des heures ou des jours.
        if TobAction.GetValue ('RAC_DELAIRAPPEL') = '' then
           TobAction.PutValue('RAC_DATERAPPEL',DateAction)
        else
            if TobAction.GetValue ('RAC_DELAIRAPPEL') < '024' then
                TobAction.PutValue('RAC_DATERAPPEL',DateAction-EncodeTime(TobAction.GetValue ('RAC_DELAIRAPPEL'), 0, 0, 0))
            else
                TobAction.PutValue('RAC_DATERAPPEL',PlusDate(DateAction, (TobAction.GetValue ('RAC_DELAIRAPPEL')/24)* (-1),'J'));
        end;
    if ( ctxAffaire in V_PGI.PGIContexte) or ( ctxGCAFF in V_PGI.PGIContexte) then
      begin
      TobAction.PutValue ('RAC_AFFAIRE',GetField('RCH_AFFAIRE'));
      TobAction.PutValue ('RAC_AFFAIRE0',GetField('RCH_AFFAIRE0'));
      TobAction.PutValue ('RAC_AFFAIRE1',GetField('RCH_AFFAIRE1'));
      TobAction.PutValue ('RAC_AFFAIRE2',GetField('RCH_AFFAIRE2'));
      TobAction.PutValue ('RAC_AFFAIRE3',GetField('RCH_AFFAIRE3'));
      TobAction.PutValue ('RAC_AVENANT',GetField('RCH_AVENANT'));
      end;
    if i = 0 then
      begin
        //TOBCopieChamp(TobAction,TobActionSav);
        TobAction.PutValue ('RAC_INTERVINT','');
      end;
    TobAction.InsertDB(Nil);
    RTCreationLiensActions (TobAction);
    CreatInfosCompl(TobActionsChainage,GetField('RCH_AUXILIAIRE'),nbact);

    if EnvoiMail = 'X' then
     begin
     if MailResp = RESPONSABLE then
        begin
        if TobAction.GetValue ('RAC_INTERVENANT') <> '' then
           ActEnvoiMessage ('R',TobAction,MailAuto)
        else
           if (not V_Pgi.SilentMode) then PGIBox('Envoi de message impossible : absence de responsable','Chaînage d''actions');
        end
     else
        if TobAction.GetValue ('RAC_NUMEROCONTACT') <> 0 then
           ActEnvoiMessage ('I',TobAction,MailAuto)
        else
           if (not V_Pgi.SilentMode) then PGIBox('Envoi de message impossible : absence d''interlocuteur','Chaînage d''actions');
     end;
    if (OutLook = 'X') and (TobAction.GetValue('RAC_INTERVENANT')=VH_RT.RTResponsable) then
     begin
     Duree:=trunc(TobAction.GetValue('RAC_DUREEACTION'));
     AddRDV(TobAction.GetValue ('RAC_LIBELLE'),RechDom('RTTIERSPRO',GetField ('RCH_TIERS'),False),TobAction.GetValue ('RAC_BLOCNOTE'),TobAction.GetValue ('RAC_DATEACTION'),TobAction.GetValue ('RAC_HEUREACTION') ,Duree) ;
     end;
    FreeAndNil(TobAction);
    TobActionsChainage:=TobActionsLocal.FindNext(['RPA_CHAINAGE','RPA_PRODUITPGI'],[GetField('RCH_CHAINAGE'),stProduitpgi],false) ;
    inc(i);
    end;
  FreeAndNil(TobActionsLocal);
  //FreeAndNil(TobActionSav);
end;

procedure TOM_ACTIONSCHAINEES.TOBCopieChamp(FromTOB, ToTOB : TOB);
var i_pos,i_ind1,indiceDeb,indiceFin: integer;
    FieldNameTo,FieldNameFrom,St:string;
    PrefixeTo : string;
begin
  if FromTOB.Numtable > 0 then // reelle
    begin
    indiceDeb := 1;
    indiceFin := FromTOB.NbChamps;
    end
  else
    begin
    indiceDeb := 1000;
    indiceFin := 1000 + Pred(FromTOB.ChampsSup.Count);
    end;

  if ToTOB.Numtable > 0 then // reelle
    PrefixeTo := TableToPrefixe (ToTOB.NomTable)
  else
    PrefixeTo := ToTOB.NomTable; //virtuelle

  for i_ind1 := indiceDeb to indiceFin do
    begin
      FieldNameFrom := FromTOB.GetNomChamp(i_ind1);
      St := FieldNameFrom ;
      i_pos := Pos ('_', St) ;
      Delete (St, 1, i_pos-1) ;
      FieldNameTo := PrefixeTo + St ;
      if ToTOB.FieldExists(FieldNameTo) and (St <> '_BLOCNOTE') then
        ToTOB.PutValue(FieldNameTo, FromTOB.GetValue(FieldNameFrom))
    end;
end;

Procedure TOM_ACTIONSCHAINEES.CreatInfosCompl(TobActionsChainage : Tob; StAuxi : String; INumAct : Integer);
var ExisteInfos : boolean;
    FieldNameFrom,FieldNameTo : string;
    i : integer;
    TobInfos : Tob;
begin
  if soRtgestinfos001 = False then exit;

  ExisteInfos:=False;
  for i :=1 to TobActionsChainage.NbChamps do
  begin
    FieldNameFrom:= TobActionsChainage.GetNomChamp(i);
    if (copy(FieldNameFrom,11,5) = 'TEXTE') and
        (TobActionsChainage.GetValue(FieldNameFrom)<>'') then ExisteInfos:=True;
    if (copy(FieldNameFrom,11,3) = 'VAL') and
        (TobActionsChainage.GetValue(FieldNameFrom)<>0) then ExisteInfos:=True;
    if (copy(FieldNameFrom,11,5) = 'TABLE') and
        (TobActionsChainage.GetValue(FieldNameFrom)<>'') then ExisteInfos:=True;
    if (copy(FieldNameFrom,11,3) = 'MUL') and
        (TobActionsChainage.GetValue(FieldNameFrom)<>'') then ExisteInfos:=True;
    if (copy(FieldNameFrom,11,4) = 'BOOL') and
        (TobActionsChainage.GetValue(FieldNameFrom)<>'-') then ExisteInfos:=True;
    if (copy(FieldNameFrom,11,4) = 'DATE') and
        (TobActionsChainage.GetValue(FieldNameFrom)<>iDate1900 ) then ExisteInfos:=True;
    if (copy(FieldNameFrom,4,8) = 'BLOCNOTE') and
        (TobActionsChainage.GetValue(FieldNameFrom)<>'') then ExisteInfos:=True;
    if ExisteInfos=True then break;
  end;
  if not ExisteInfos then exit;
  TobInfos:=TOB.create ('RTINFOS001',NIL,-1);
  TobInfos.InitValeurs ;
  TobInfos.PutValue ('RD1_CLEDATA',StAuxi+';'+IntToStr(INumAct));
  for i :=1 to TobInfos.NbChamps do
  begin
    FieldNameTo:= TobInfos.GetNomChamp(i);
    if (copy(FieldNameTo,1,7) <> 'RD1_RD1') and
       (copy(FieldNameTo,1,12) <> 'RD1_BLOCNOTE') then continue;
    if (copy(FieldNameTo,1,7) = 'RD1_RD1') then
      FieldNameFrom  := 'RPA_RPR' + copy(FieldNameTo,8,length(FieldNameTo))
    else
      FieldNameFrom  := 'RPA_BLOCNOTE';
    TobInfos.PutValue(FieldNameTo, TobActionsChainage.GetValue(FieldNameFrom));
  end;
  TobInfos.InsertDB(Nil);
  FreeAndNil(TobInfos);
end;

procedure AGLRTChainageAttacheAction( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_ACTIONSCHAINEES) then TOM_ACTIONSCHAINEES(OM).RTChainageAttacheAction else exit;
end;

procedure AGLRTChainagenouvellePiece( parms: array of variant; nb: integer )  ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_ACTIONSCHAINEES) then TOM_ACTIONSCHAINEES(OM).ChainageNouvellePiece;
end;


procedure TOM_ACTIONSCHAINEES.ActEnvoiMessage (TypeMessage:String; TobActionSav:Tob;MailAuto:String);
var StDestinataire,StSQL,StWhere,StOr,Critere,Objet: string;
    IContact,i : integer;
    Q : TQuery;
    TobContact,TobIntervint:TOB;
//    BODY:Tstrings;
    Memo : TMEMO;
    Liste :HTStringList;
begin
TobContact := Nil;
if (TypeMessage = 'I') then
    begin
    stWhere:='';
    StOr:='';
    StDestinataire:=TobActionSav.GetValue ('RAC_DESTMAIL');
    if StDestinataire <> TraduireMemoire('<<Tous>>') then
       begin
       if (StDestinataire <> '') or (TobActionSav.GetValue ('RAC_NUMEROCONTACT') <> 0) then
          begin
          While StDestinataire <> '' do
              begin
              i :=ReadTokenI(StDestinataire);
              if i <> 0 then     // car chaîne commence par ; voir modif 0602
                begin
                StWhere := StWhere + StOr + 'C_NUMEROCONTACT='+ IntToStr(i);
                StOr := ' OR ';
                end
              end;
          if TobActionSav.GetValue ('RAC_NUMEROCONTACT') <> 0 then
             begin
             StWhere := StWhere + StOr + 'C_NUMEROCONTACT='+ IntToStr(TobActionSav.GetValue('RAC_NUMEROCONTACT'));
             end;
          end else
          begin
          StWhere := 'C_NUMEROCONTACT =' + IntToStr(0);
          end;
       end;
    StSQL:='SELECT C_NUMEROCONTACT,C_NOM,C_FONCTION,C_TELEPHONE,C_RVA'+
               ',C_CIVILITE FROM CONTACT WHERE C_TYPECONTACT = "T" AND C_AUXILIAIRE = "'+TobActionSav.GetValue('RAC_AUXILIAIRE')+'"';
    if StWhere<>'' then StSQL:=StSQL+' AND ('+stWhere+')';

    TobContact:=TOB.create ('les contacts',NIL,-1);
    Q := OpenSql (StSQL,TRUE);
    TobContact.LoadDetailDB ('CONTACT', '', '', Q, False);
    ferme(Q) ;
    end;

// cd 010301 : Remplacement table Commercial par table Ressource
stWhere:='';
StOr:='';
StDestinataire:=TobActionSav.GetValue ('RAC_INTERVINT');
if StDestinataire <> TraduireMemoire('<<Tous>>') then
   begin
   Stwhere := 'WHERE ';
   if StDestinataire <> '' then
      begin
      While StDestinataire <> '' do
          begin
          Critere :=ReadTokenSt(StDestinataire);
          StWhere := StWhere + StOr + 'ARS_RESSOURCE= "'+ Critere+'"';
          StOr := ' OR ';
          end;
      end else
      begin
      StWhere := 'WHERE ARS_RESSOURCE = "#######"';
      end;
   end;

if (TypeMessage <> 'I') and (TobActionSav.GetValue ('RAC_INTERVENANT') <> '') then
   begin
   StWhere := StWhere + 'or ARS_RESSOURCE= "'+ TobActionSav.GetValue('RAC_INTERVENANT')+'"';
   end;

StSQL:='SELECT ARS_RESSOURCE,ARS_LIBELLE,ARS_TYPERESSOURCE,ARS_EMAIL'+
           ' FROM RESSOURCE ';
if StWhere<>'' then StSQL:=StSQL+stWhere;

TobIntervint:=TOB.create ('les intervenants',NIL,-1);
Q := OpenSql (StSQL,TRUE);
TobIntervint.LoadDetailDB ('RESSOURCE', '', '', Q, False);
ferme(Q) ;

StDestinataire := '';
StWhere := '';
if (TypeMessage = 'I') then
begin
for i:=0 to TobContact.detail.count-1 do
    begin
    if TobContact.detail[i].GetValue('C_RVA') <> '' then
       begin
       IContact:=TobContact.detail[i].GetValue('C_NUMEROCONTACT');
       if IContact = TobActionSav.GetValue('RAC_NUMEROCONTACT') then
          begin
          StWhere := TobContact.detail[i].GetValue('C_RVA');
          end
       else
           begin
           StDestinataire := StDestinataire + TobContact.detail[i].GetValue('C_RVA') +';';
           end;
       end;
    end;
end;

for i:=0 to TobIntervint.detail.count-1 do
    begin
    if TobIntervint.detail[i].GetValue('ARS_EMAIL') <> '' then
       begin
       critere := TobIntervint.detail[i].GetValue('ARS_RESSOURCE');
       if (critere = TobActionSav.GetValue('RAC_INTERVENANT')) and (TypeMessage <> 'I') then
          begin
          StWhere := TobIntervint.detail[i].GetValue('ARS_EMAIL');
          end
       else
          begin
          StDestinataire := StDestinataire + TobIntervint.detail[i].GetValue('ARS_EMAIL') +';';
          end;
       end;
    end;

if StWhere <> '' then
   begin
   Objet := '';

   if Generation_LibTiers <> '' then Objet := TraduireMemoire('Client')+' : '+Generation_LibTiers
   else if GetControlText('TRCH_TIERS_') <> '' then
      Objet := TraduireMemoire('Client')+' : '+GetControlText('TRCH_TIERS_');

   if soRtmesstypeact = True then
      begin
      if Objet <> '' then Objet := Objet + ' - ' + Getfield('RCH_LIBELLE')
      else Objet := Getfield('RCH_LIBELLE');
      end;

   //Objet := RechDom('RTTYPEACTION',GetField('RAC_TYPEACTION'),FALSE);
   if soRtmesslibact = True  then
      begin
      if Objet <> '' then Objet := Objet + ' - ' + TobActionSav.GetValue('RAC_LIBELLE')
      else Objet := TobActionSav.GetValue('RAC_LIBELLE')
      end;

   if (soRtmessbl = True) and Assigned(ecran) then
      begin
      Memo := TMemo(GetControl('RCH_BLOCNOTE'));
      Liste:=HTStringList.Create ;
      Liste.Text := Memo.text;
      PGIEnvoiMail (Objet,StWhere,StDestinataire,Liste,'',MailAuto='X',1,'','');
      Liste.free;
      end
      else
      PGIEnvoiMail (Objet,StWhere,StDestinataire,Nil,'',MailAuto='X',1,'','');

   end
else
   begin
   if (not V_Pgi.SilentMode) then 
   PGIBox('Envoi de message impossible : pas d''adresse e-mail','Chaînage d''actions');
   end;
if TypeMessage = 'I' then FreeAndNil(TobContact);
FreeAndNil(TobIntervint);
end;
procedure AGLRTGenereLotActions( parms: array of variant; nb: integer )  ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_ACTIONSCHAINEES) then TOM_ACTIONSCHAINEES(OM).RTGenereLotActions;
end;

function Chainage_MyAfterImport (Sender: TObject) : string;
var  OM : TOM ;
begin
result := '';
if sender is TFFICHE then OM := TFFICHE(Sender).OM else exit;
if (OM is TOM_ACTIONSCHAINEES) then result := TOM_ACTIONSCHAINEES(OM).GetInfoGed else exit;
end;

procedure Chainage_GestionBoutonGED (Sender: TObject);
var  OM : TOM ;
begin
if sender is TFFICHE then OM := TFFICHE(Sender).OM else exit;
if (OM is TOM_ACTIONSCHAINEES) then TOM_ACTIONSCHAINEES(OM).GestionBoutonGED else exit;
end;

function TOM_ACTIONSCHAINEES.GetInfoGed : string;
begin
if (not(DS.State in [dsInsert])) and (stProduitpgi = 'GRC') then result := 'Tiers='+GetField('RCH_TIERS')+';'+'Chainage='+intTostr(GetField('RCH_NUMERO'));
end;

procedure TOM_ACTIONSCHAINEES.RTAppelGED_OnClick (Sender: TObject) ;
begin
if (ds<>nil) and not(DS.State in [dsInsert]) then
   begin
   AGLLanceFiche('RT','RTRECHDOCGED','RTD_TIERS='+GetField ('RCH_TIERS')+';RTD_NUMCHAINAGE='+IntToStr(GetField ('RCH_NUMERO')),'','Objet=CHA;Tiers='+GetField ('RCH_TIERS')+';Chainage='+IntToStr(GetField ('RCH_NUMERO')));
   GestionBoutonGED;
   end;
end;

{ GC/GRC : MNG / gestion des alertes }
procedure TOM_ACTIONSCHAINEES.ListAlerte_OnClick_Rch(Sender: TObject);
begin
if (GetField('RCH_NUMERO') <> 0) and(AlerteActive('RCH')) then
   AGLLanceFiche('Y','YALERTES_MUL','YAL_PREFIXE=RCH','','ACTION=CREATION;MONOFICHE;CHAMP=RCH_NUMERO;VALEUR='
      +IntToStr(GetField('RCH_NUMERO'))+';LIBELLE='+GetField('RCH_LIBELLE')) ;
end ;

procedure TOM_ACTIONSCHAINEES.Alerte_OnClick_Rch(Sender: TObject);
begin
if (GetField('RCH_NUMERO') <> 0) and(AlerteActive('RCH')) then
   AGLLanceFiche('Y','YALERTES','','','ACTION=CREATION;MONOFICHE;CHAMP=RCH_NUMERO;VALEUR='
   +IntToStr(GetField('RCH_NUMERO'))+';LIBELLE='+GetField('RCH_LIBELLE')) ;
VH_EntPgi.TobAlertes.ClearDetail;
end;


Procedure TOM_ACTIONSCHAINEES.GestionBoutonGED;
BEGIN
if (GetParamSocSecur('SO_RTGESTIONGED',False)) and (ds<>nil) and not(DS.State in [dsInsert]) and (ExisteSQL('SELECT RTD_DOCID FROM RTDOCUMENT WHERE RTD_TIERS="'+GetField('RCH_TIERS')+'" AND RTD_NUMCHAINAGE='+IntToStr(GetField ('RCH_NUMERO')))) then
     SetControlVisible ('BDOCGEDEXIST',True)
else SetControlVisible ('BDOCGEDEXIST',False);
END;

Initialization
registerclasses([TOM_ACTIONSCHAINEES]) ;
RegisterAglProc('RTChainageAttacheAction', TRUE , 0, AGLRTChainageAttacheAction);
RegisterAglProc('RTChainageNouvellePiece', True,0,AGLRTChainageNouvellePiece) ;
RegisterAglProc('RTGenereLotActions', True,0,AGLRTGenereLotActions) ;
end.
