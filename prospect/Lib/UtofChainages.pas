{***********UNITE*************************************************
Auteur  ...... : MNG
Créé le ...... : 30/10/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CHAINAGES ()
Mots clefs ... : TOF;CHAINAGES
*****************************************************************}
Unit UtofChainages ;

Interface

Uses Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     FE_Main,db,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ELSE}
     Maineagl,
{$ENDIF}
{$IFDEF GIGI}
      EntGc,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     uTableFiltre,
     SaisieList,
     hTB97,
     UtilRT,M3FP,ParamSoc,UTob, EntRT, UTom, UtomActionchaine, HRichOle,extctrls, windows,
     UtilSelection,hqry,menus;
Const
    ANNULEE  : String = 'ANU';
    REALISEE : String = 'REA';
    StVisee : String = 'Chaînage terminé, pièce(s) visée(s) : ';
Type
  TOF_CHAINAGES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure DoClick( Sender: TObject );
    procedure DoDblClick( Sender: TObject );
    procedure DoSetNavigate( Sender: TObject );
    procedure OnClose              ; override;
  Private
    TF: TTableFiltre;
    stProduitpgi : string;
    soRtchapplicrit,GestionIntervention,OrigineParc,OrigineTiers : boolean;
    procedure RTChainageCherche;
    procedure RTMenuRechTree;
    function DoTestAndMajInterv (stChainage : String) : String;
    procedure DoSetLibelleLibreInterv;
  end ;

Function RTLanceFiche_Chainages(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

implementation

//GP_20080201_TS_GC15341 >>>
uses
  FactUtil
  ,FactTob
  ;
//GP_20080201_TS_GC15341 <<<

Function RTLanceFiche_Chainages(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_CHAINAGES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CHAINAGES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CHAINAGES.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CHAINAGES.OnLoad ;
var xx_where,StWhere,DebutWhere,Confid: string;
    DateDeb,DateFin : TDateTime;
    i: integer;
begin
  Inherited ;
  StWhere := '';
  //SetControlText('XX_WHERE',StWhere) ;
  if stProduitpgi = 'GRC' then Confid:='CON' else Confid:='CONF';

  xx_where:=RTXXWhereConfident(Confid);
  SetControlText('XX_WHERE',xx_where) ;
  if GetControlText('TOUTEREALISEES') = 'X' then
  begin
    StWhere:='NOT EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_NUMCHAINAGE = RCH_NUMERO';
    StWhere:=StWhere + ' and RAC_PRODUITPGI="'+stProduitpgi+'" AND (RAC_ETATACTION="PRE" or RAC_ETATACTION="NRE")))';
  end
  else
  begin
    DebutWhere:='EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_NUMCHAINAGE = RCH_NUMERO';
    DebutWhere:=DebutWhere+' and RAC_PRODUITPGI="'+stProduitpgi+'"';

    DateDeb := StrToDate(GetControlText('DATEACTION'));
    DateFin := StrToDate(GetControlText('DATEACTION_'));
    if (GetControlText('TYPEACTION') <> '') and (GetControlText('TYPEACTION') <> TraduireMemoire('<<Tous>>')) then
    begin
      if StWhere='' then StWhere:=DebutWhere;
      StWhere:=StWhere + ' AND RAC_TYPEACTION="'+GetControlText('TYPEACTION')+'"';
    end;
    if (GetControlText('ETATACTION') <> '') and (GetControlText('ETATACTION') <> TraduireMemoire('<<Tous>>')) then
    begin
      if StWhere='' then StWhere:=DebutWhere;
      StWhere:=StWhere + ' AND RAC_ETATACTION="'+GetControlText('ETATACTION')+'"';
    end;
    if GetControlText('RESPONSABLE') <> '' then
    begin
      if StWhere='' then StWhere:=DebutWhere;
      StWhere:=StWhere + ' AND RAC_INTERVENANT="'+GetControlText('RESPONSABLE')+'"';
      end;
    for i:=1 to 3 do
      begin
      if (GetControlText('TABLELIBRE'+intToStr(i)) <> '') and (GetControlText('TABLELIBRE'+intToStr(i)) <> TraduireMemoire('<<Tous>>')) then
          begin
          if StWhere='' then StWhere:=DebutWhere;
          StWhere:=StWhere + ' AND RAC_TABLELIBRE'+intToStr(i)+'="'+GetControlText('TABLELIBRE'+intToStr(i))+'"';
          end;
      end;
    if ( DateDeb <> iDate1900 ) or (DateFin <> iDate2099 ) then
    begin
      if StWhere='' then StWhere:=DebutWhere;
      StWhere := StWhere + ' AND RAC_DATEACTION >= "'+UsDateTime(DateDeb) +'" AND RAC_DATEACTION <= "'+UsDateTime(DateFin)+'"';
    end;
    if StWhere<>'' then StWhere:=StWhere+'))';

    // Uniquement les responsable avec action prévue
    if (GetControlText('UNIQUE') = 'X') and (GetControlText('RESPONSABLE')<>'')
       and (GetControlText('ETATACTION')<>'') then
    begin
      StWhere := StWhere + ') AND ((SELECT count(*) FROM ACTIONS WHERE (RAC_NUMCHAINAGE = RCH_NUMERO ';
      StWhere := StWhere + ' and RAC_PRODUITPGI="'+stProduitpgi+'" ';
      StWhere := StWhere + 'AND RAC_ETATACTION="'+GetControlText('ETATACTION')+'")) = 1 ';
    end;
  end;

  if (GetControlText('XX_WHERE') = '') then
    SetControlText('XX_WHERE',StWhere)
  else
     begin
     xx_where := GetControlText('XX_WHERE');
     if StWhere <> '' then
     begin
       xx_where := xx_where + ' and (' + StWhere + ')';
       SetControlText('XX_WHERE',xx_where) ;
     end;
     end;

end ;

procedure TOF_CHAINAGES.OnArgument (S : String ) ;
var F : TForm;
    i : integer;
    Confid,stNum,ori : string;
begin
  Inherited ;
  OrigineParc:=false;OrigineTiers:=false;
  stProduitpgi:=ReadToKenSt(S);
  ori:=ReadToKenSt(S);
  if (ori='ORIGINEPARC') or (ori='ORIGINETIERS') then
    begin
    if (ori='ORIGINEPARC') then
      OrigineParc:=true;
    stNum:=ReadToKenSt(S);
    stNum:=copy(stNum,pos('=',stNum)+1,length(stNum));
    if ValeurI(stNum) <> 0 then
      SetControlText('RCH_NUMERO',stNum);
    SetControlEnabled('RCH_NUMERO',false);
    end;
  GestionIntervention:=false;
  if (Ecran<>nil) and (Ecran is TFSaisieList ) then
  begin
    if copy(ecran.name,1,17) = 'WINTERVENTION_FSL' then GestionIntervention:=true;
    TF := TFSaisieList(Ecran).LeFiltre;
    TF.OnSetNavigate := DoSetNavigate;
    if stProduitpgi='GRF' then
      begin
      TF.LaTreeListe:='RFTVCHAINAGE';
      TF.LaGridListe:='RFMULACTIONSCH';
      for i:=1 to 9 do
         begin
         SetControlVisible ('YTC_TABLELIBRETIERS'+IntToStr(i),false);
         SetControlVisible ('TYTC_TABLELIBRETIERS'+IntToStr(i),false);
         end;
      SetControlVisible ('YTC_TABLELIBRETIERSA',false);
      SetControlVisible ('TYTC_TABLELIBRETIERSA',false);
      SetControlVisible ('T_NATUREAUXI',false);
      SetControlVisible ('TT_NATUREAUXI',false);
      SetControlVisible ('T_REPRESENTANT',false);
      SetControlVisible ('TT_REPRESENTANT',false);
      for i:=1 to 3 do
         begin
         SetControlVisible ('TABLELIBRE'+IntToStr(i),false);
         SetControlVisible ('TRAC_TABLELIBRE'+IntToStr(i)+'_',false);
         end;
      for i:=1 to 3 do
         begin
         SetControlVisible ('RCH_TABLELIBRECH'+IntToStr(i),false);
         SetControlVisible ('TRCH_TABLELIBRECH'+IntToStr(i),false);
         end;
      if Assigned(GetControl('POPCOMPLEMENT')) then
        begin
        TPopupMenu(GetControl('POPCOMPLEMENT')).items[2].visible:=false; { projet }
        TPopupMenu(GetControl('POPCOMPLEMENT')).items[4].visible:=false; { infos compl. }
        end;
      if Assigned(GetControl('RCH_CHAINAGE')) then
        SetControlProperty ('RCH_CHAINAGE', 'Plus', 'RPG_PRODUITPGI="GRF"');
      end
    else
      begin
      for i:=1 to 3 do
         begin
         SetControlVisible ('YTC_TABLELIBREFOU'+IntToStr(i),false);
         SetControlVisible ('TYTC_TABLELIBREFOU'+IntToStr(i),false);
         end;
      for i:=1 to 3 do
         begin
         SetControlVisible ('TABLELIBREF'+IntToStr(i),false);
         SetControlVisible ('TRAC_TABLELIBREF'+IntToStr(i)+'_',false);
         end;
      for i:=1 to 3 do
         begin
         SetControlVisible ('RCH_TABLELIBRECHF'+IntToStr(i),false);
         SetControlVisible ('TRCH_TABLELIBRECHF'+IntToStr(i),false);
         end;
      if Assigned(GetControl('RCH_CHAINAGE')) then
        SetControlProperty ('RCH_CHAINAGE', 'Plus', 'RPG_PRODUITPGI="GRC"');
      if Assigned(GetControl('WIV_CHAINAGE')) then
        SetControlProperty ('WIV_CHAINAGE', 'Plus', 'RPG_PRODUITPGI="GRC"');
      end;
  end;
  if Assigned(GetControl('TreeEntete')) then
    THTreeView(GetControl( 'TreeEntete' )).OnDblClick := DoDblClick;
  if Assigned(GetControl('BNEWCHAINAGE')) then
    TToolBarButton97(GetControl( 'BNEWCHAINAGE' )).OnClick := DoClick;
  F := TForm (Ecran);
  if stProduitpgi <> 'GRF' then
    begin
    Confid:='CON' ;
    soRtchapplicrit:=GetParamSocSecur('SO_RTCHAPPLICRIT',false);
    MulCreerPagesCL(F,'NOMFIC=GCTIERS');
    end
  else
    begin
    Confid:='CONF';
    soRtchapplicrit:=GetParamSocSecur('SO_RFCHAPPLICRIT',false);
    if GetParamSocSecur('SO_RTGESTINFOS003',false) = True then
        MulCreerPagesCL(F,'NOMFIC=GCFOURNISSEURS');
    end;

  SetControlText('XX_WHERE',RTXXWhereConfident(Confid));
  TF.WhereEntete := RecupWhereCritere( TFSaisieList(ecran).Pages );
{$Ifdef GIGI}
 if (GetControl('RAC_OPERATION') <> nil) and Not GetParamSocSecur('SO_AFRTOPERATIONS',false)
    then  begin
    SetControlVisible('RAC_OPERATION',false);
    SetControlVisible('TRAC_OPERATION',false);
    end;
 if (GetControl('RAC_PROJET') <> nil) and Not GetParamSocSecur('SO_RTPROJGESTION',false)
    then  begin
    SetControlVisible('RAC_PROJET',false);
    SetControlVisible('TRAC_PROJET',false);
    end;
 If (Not VH_GC.GaSeria) or not (GetParamSocSecur ('SO_AFRTPROPOS',false)) then
   begin
   if (GetControl('TRAC_PERSPECTIVE') <> nil) then  SetControlVisible('TRAC_PERSPECTIVE',false);
   if (GetControl('NUM_PERSPECTIVE') <> nil) then  SetControlVisible('NUM_PERSPECTIVE',false);
   if (GetControl('TRAC_AFFAIRE0') <> nil) then  SetControlVisible('TRAC_AFFAIRE0',false);
   if (GetControl('RAC_AFFAIRE1') <> nil) then  SetControlVisible('RAC_AFFAIRE1',false);
   if (GetControl('RAC_AFFAIRE2') <> nil) then  SetControlVisible('RAC_AFFAIRE2',false);
   if (GetControl('RAC_AFFAIRE3') <> nil) then  SetControlVisible('RAC_AFFAIRE3',false);
   if (GetControl('RAC_AVENANT') <> nil) then  SetControlVisible('RAC_AVENANT',false);
   if (GetControl('BEFFACEAFF1') <> nil) then  SetControlVisible('BEFFACEAFF1',false);
   if (GetControl('BSELECTAFF1') <> nil) then  SetControlVisible('BSELECTAFF1',false);
   end;
 if (GetControl('T_ZONECOM') <> nil) then SetControlVisible('T_ZONECOM',false);
 if (GetControl('TT_REPRESENTANT') <> nil) then SetControlVisible('TT_REPRESENTANT',false);
 if (GetControl('T_REPRESENTANT') <> nil) then SetControlVisible('T_REPRESENTANT',false);
{$endif}
  if GestionIntervention then DoSetLibelleLibreInterv;
end ;


procedure TOF_CHAINAGES.DoClick( Sender: TObject );
var Ret,NumChainage : String;
begin
  if UpperCase(TControl( Sender ).name) = 'BNEWCHAINAGE' then
    begin
    if not GestionIntervention then
      begin
      Ret := AGLLanceFiche('RT','RTACTIONSCHAINE','','','MONOFICHE;ACTION=CREATION;LISTEMENU;PRODUITPGI='+stProduitpgi);
      NumChainage:=ReadToKenSt(Ret);
      end
    else
      NumChainage := AGLLanceFiche('W','WINTERVENTION','','','MONOFICHE;ACTION=CREATION;LISTEMENU;ORIGINEPARC;PRODUITPGI='+stProduitpgi) ;
    if NumChainage <> '' then
      begin
      if (TF.IndexEntete=0) then  // pour savoir si on était sans entête
        TF.RefreshEntete( '' )  // la ça marche mais le filtre de sélection et réappliqué donc si ta fiche n'y est pas elle n'apparaît pas
      else
        if GestionIntervention then
          TF.RefreshEntete( NumChainage )
        else
          TF.RefreshEntete( NumChainage+';'+stProduitpgi );
      end;
    end;
end;
procedure TOF_CHAINAGES.DoSetNavigate( Sender: TObject );
var i : integer;
  ModChainage,NumChainage,StPiecesVisees,Requete,rch_intervenant,rch_libelle,stInterv : String;
  Termine,rch_termine : boolean;
  TobRCH,TobTypeEncours,TobPiece : tob;
  OM : TOM;
  //St : TStrings;
  St : HTStringList;
  Q,QQ : TQuery;
  Memo : TRichEdit;
  ThePanel:TPanel;
begin
  St:=Nil;
  ThePanel:=Nil;
  TobTypeEncours:=nil;
  TobRCH:=nil;
  rch_termine:=false;
  rch_intervenant:='';
  rch_libelle:='';

  if not GestionIntervention then
    begin
    ModChainage := TF.TobFiltre.GetString('RCH_CHAINAGE');
    NumChainage:=TF.TobFiltre.GetString('RCH_NUMERO');
    rch_termine:=TF.TobFiltre.GetBoolean('RCH_TERMINE');
    rch_intervenant:=TF.TobFiltre.GetString('RCH_INTERVENANT');
    rch_libelle:=TF.TobFiltre.GetString('RCH_LIBELLE');
    end
  else
    begin
    ModChainage := '';
    ModChainage := TF.TobFiltre.GetString( 'WIV_CHAINAGE' );
    NumChainage:=TF.TobFiltre.GetString('WIV_NUMCHAINAGE');
    {if NumChainage = '0' then
      SetControlEnabled ('Binsert',false)
    else
      SetControlEnabled ('Binsert',true);}
    if ( TF.State = dsBrowse) and (NumChainage <> '0') then
      begin
      TobRCH:= TOB.Create('ACTIONSCHAINEES', nil, -1);
      TobRCH.SelectDB(NumChainage+';"GRC"',Nil );
      if Assigned(TobRCH) then
        begin
        rch_termine:=TobRCH.GetBoolean('RCH_TERMINE');
        rch_intervenant:=TobRCH.GetString('RCH_INTERVENANT');
        rch_libelle:=TobRCH.GetString('RCH_LIBELLE');
        end;
      end;
    end;

  { sur option du type chainage, on met à jour le status Termine si toutes les
    actions du chainage sont réalisées ou abandonnées }
  if ModChainage <> '' then
  begin
    VH_RT.TobTypesChainage.Load;
    TobTypeEncours:=VH_RT.TobTypesChainage.FindFirst(['RPG_CHAINAGE','RPG_PRODUITPGI'],[ModChainage,stProduitpgi],TRUE) ;
  end;
  if (TobTypeEncours <> Nil) and ( TF.State = dsBrowse) then
    if (TobTypeEncours.GetValue('RPG_FERMEAUTO')= 'X' ) and
        (not rch_termine) then
    begin
      Termine:=true;
      if TF.TOBFiltre.detail.count = 0 then Termine:=false;
      for i:=0 to TF.TOBFiltre.detail.count-1 do
      begin
        if ( TF.TOBFiltre.detail[i].GetValue('RAC_ETATACTION') <> REALISEE ) and
           ( TF.TOBFiltre.detail[i].GetValue('RAC_ETATACTION') <> ANNULEE ) then
        begin
          Termine:=false;
          break;
        end;
      end;
      if Termine then
      begin
        StPiecesVisees:='';
        if ( TobTypeEncours.GetValue('RPG_VISA') = 'X' ) then
          begin
          StPiecesVisees:='';
          TobPiece:=tob.create('_PIECE',Nil,-1) ;
          Requete:='SELECT GP_NATUREPIECEG,GP_NUMERO,GP_SOUCHE,GP_INDICEG,GP_ETATVISA from CHAINAGEPIECES '+
          'left join PIECE on GP_NUMERO=RLC_NUMERO and GP_SOUCHE=RLC_SOUCHE and GP_INDICEG=RLC_INDICEG '+
          'and GP_NATUREPIECEG=RLC_NATUREPIECEG'+
          ' where RLC_NUMCHAINAGE='+NumChainage+' and RLC_PRODUITPGI="'+stProduitpgi+'"';

          QQ:=OpenSQL(Requete,True);
          TobPiece.LoadDetailDB('PIECE','','',QQ,false,true) ;
          ferme(QQ);
          for i:=0 to TobPiece.detail.count-1 do
            begin
            if ( TobPiece.detail[i].GetValue('GP_ETATVISA') = 'ATT' ) then
              begin
              ExecuteSQl('UPDATE PIECE set GP_ETATVISA="VIS" Where'
                +' GP_NATUREPIECEG="'+TobPiece.detail[i].GetValue('GP_NATUREPIECEG')
                +'" AND GP_SOUCHE="'+TobPiece.detail[i].GetValue('GP_SOUCHE')
                +'" AND GP_NUMERO='+IntToStr(TobPiece.detail[i].GetValue('GP_NUMERO'))
                +' AND GP_INDICEG='+IntToStr(TobPiece.detail[i].GetValue('GP_INDICEG')));

//GP_20080201_TS_GC15341 >>>
              { Action après modification du "Visa" d'une pièce}
              GPDoActionAfterUpdatePiece(Tob2CleDoc(TobPiece.detail[i]));
//GP_20080201_TS_GC15341 <<<

              if StPiecesVisees = '' then
                 StPiecesVisees:=StVisee;
              StPiecesVisees:=StPiecesVisees+#13#10+'  '+RechDom('GCNATUREPIECEG',TobPiece.detail[i].GetValue('GP_NATUREPIECEG'),False)+' n° : '+IntToStr(TobPiece.detail[i].GetValue('GP_NUMERO'))
              end;
            end;
          TobPiece.free ;
          if StPiecesVisees<>'' then StPiecesVisees:=StPiecesVisees+#13#10;
          end;
        if ( TobTypeEncours.GetValue('RPG_MAILTERMINE') = 'X' ) and
           ( rch_intervenant <> '' ) and { envoi mail si intervenant renseigné }
           ( rch_intervenant <> VH_RT.RTResponsable ) then { uniquement si le responsable n'est pas l'utilisateur }
          begin
          Q := OpenSQL('SELECT RCH_BLOCNOTE FROM ACTIONSCHAINEES WHERE RCH_NUMERO='+
               NumChainage+' and RCH_PRODUITPGI="'+stProduitpgi+'"', True);
          if not Q.Eof then
            begin
            ThePanel := TPanel.Create( nil );
            ThePanel.Visible := False;
            ThePanel.ParentWindow := GetDesktopWindow;
            Memo := TRichEdit.Create( ThePanel );
            Memo.Parent:=ThePanel;
            Memo.width:=525;
            StringTorich(Memo,Q.Findfield('RCH_BLOCNOTE').AsString);
            if StPiecesVisees = '' then
              Memo.Lines.Insert(0,TraduireMemoire('Chaînage terminé :')+#13#10)
            else
              Memo.Lines.Insert(0,StPiecesVisees);

            St:=HTStringList(Memo.Lines);
            end ;
          Ferme(Q) ;

          if Assigned(St) and Assigned(GetControl('RCH_INTERVENANT')) then
            begin
            OM:=TFSaisieList(Ecran).LeFiltre.OM ;
            TOM_ACTIONSCHAINEES(OM).EnvoiMessage(rch_intervenant+'|'+rch_libelle+'|'+NumChainage,'-',TobTypeEncours.GetValue('RPG_MAILTERMAUTO'),true,St,'');
            end;
          if ThePanel <> Nil then
             ThePanel.Free;
          end;
        if not GestionIntervention then
          begin
          TF.TobFiltre.Putvalue('RCH_TERMINE','X');
          TF.TobFiltre.Putvalue('RCH_DATEFIN', Date);
          TF.TobFiltre.UpdateDB(false,false);
          if not OrigineParc then
            begin
            if stProduitpgi = 'GRC' then
              stInterv := DoTestAndMajInterv(NumChainage);
            if stInterv <> '' then
               StPiecesVisees:=StPiecesVisees+'Intervention n° : '+stInterv+' terminée.';
            end;
          end
        else
          begin
          TobRCH.Putvalue('RCH_TERMINE','X');
          TobRCH.Putvalue('RCH_DATEFIN', Date);
          TobRCH.UpdateDB(false,false);
          TF.TobFiltre.SetString('WIV_ETATINTERV','TER');
          TF.TobFiltre.Putvalue('WIV_DATEFIN', Date);
          TF.TobFiltre.UpdateDB(false,false);
          StPiecesVisees:=StPiecesVisees+'Intervention n° : '+TF.TobFiltre.GetString('WIV_IDENTIFIANT')+' terminée.';
          end;
        if StPiecesVisees <> '' then
           PgiInfo (StPiecesVisees,Ecran.Caption);
        if soRtchapplicrit = true then
           TF.RefreshEntete('');
      end;
  end;
  if (TF.State = dsBrowse) and (GestionIntervention) and (NumChainage <> '0') then
    TobRCH.free;
  if (TF.TOBFiltre.detail.count = 0) then
    begin
    SetControlEnabled('BNEWMAIL',False);
    SetControlEnabled('BCOURRIER',False);
    SetControlEnabled('BCONTACT',False);
    SetControlEnabled('BPOPOUTLOOK',False);
    SetControlEnabled('BZOOM',False);
    SetControlVisible('BDOCGEDEXIST',False);
    end
  else
    begin
    SetControlEnabled('BNEWMAIL',true);
    SetControlEnabled('BCOURRIER',true);
    SetControlEnabled('BCONTACT',true);
    SetControlEnabled('BPOPOUTLOOK',true);
    SetControlEnabled('BZOOM',true);
    end

end;

procedure TOF_CHAINAGES.DoDblClick( Sender: TObject );
var Ret,NumChainage : String;
begin
  if UpperCase(TControl( Sender ).name) = 'TREEENTETE' then
    begin
    if not GestionIntervention then
      begin
      Ret := AGLLanceFiche('RT','RTACTIONSCHAINE','',IntToStr(TF.TOBFiltre.GetValue('RCH_NUMERO'))+';'+stProduitpgi,
       'ACTION=MODIFICATION;MONOFICHE;LISTEMENU;PRODUITPGI='+stProduitpgi+';TOBACTIONS='+IntToStr(LongInt(TF.TOBFiltre)));
      NumChainage:=ReadToKenSt(Ret);
      if soRtchapplicrit = true then
         TF.RefreshEntete('')
      else
         TF.RefreshEntete(NumChainage+';'+stProduitpgi) ;
      end
    else
      begin
      AGLLanceFiche('W','WINTERVENTION','',IntToStr(TF.TOBFiltre.GetValue('WIV_IDENTIFIANT')),
       'ACTION=MODIFICATION;MONOFICHE;ORIGINEPARC');
      TF.RefreshEntete('')
      end;
  end;
end;

procedure TOF_CHAINAGES.RTChainageCherche ;
begin
OnLoad;
end ;

procedure TOF_CHAINAGES.RTMenuRechTree ;
var F : TForm;
    TS: TTabSheet;
begin
F := TForm (Ecran);
//TFSaisieList(F).Pages.ActivePageIndex:=0;
    // positionner l'Onglet "Mise en Page" en dernier
    TS:= TTabSheet(F.FindComponent('PCritere'));
    if TS <> Nil then
    begin
      TS.PageControl.ActivePage:=Nil;
//      TS.SetFocus;
      TS.PageControl.ActivePage:=TS;
    end;
end ;

procedure TOF_CHAINAGES.OnClose;
begin
  Inherited;
  if OrigineParc and not GestionIntervention then
    TFSaisieList(Ecran).Retour:=TF.TobFiltre.GetString('RCH_TERMINE');
end;

function TOF_CHAINAGES.DoTestAndMajInterv(stChainage : String) : String;
var TobInterv : tob;
    Q : TQuery;
begin
  result:='';
  Q := OpenSQL('Select * from WINTERVENTION Where WIV_NUMCHAINAGE=' + stChainage,True) ;
  if Not Q.EOF then
    begin
    TobInterv:= TOB.Create('WINTERVENTION', nil, -1);
    TobInterv.SelectDB('',Q);
    if Assigned(TobInterv) then
      begin
      TobInterv.SetString('WIV_ETATINTERV','TER');
      TobInterv.SetDateTime('WIV_DATEMODIF',date);
      TobInterv.SetString('WIV_UTILISATEUR',V_PGI.User);
      TobInterv.UpdateDB(false,false);
      Result:=TobInterv.GetString('WIV_IDENTIFIANT');
      end;
    TobInterv.free;
    end;
  Ferme(Q);
end;

procedure TOF_CHAINAGES.DoSetLibelleLibreInterv;
var iLibre : integer;
    FieldName : String;
  procedure SetVisibleWIV(FieldName: string);
  begin
    SetControlProperty('T' + FieldName, 'VISIBLE', Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
    SetControlProperty(FieldName, 'VISIBLE', Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
  end;
begin
    { Libellés champs libres }
    for iLibre := 1 to 10 do
    begin
      { tables libres }
      FieldName := 'WIV_LIBREWIV'+intToHex(iLibre, 1);
      SetControlText('T' + FieldName,RechDom('WLIBELLELIBREWIV', 'T' + FieldName, false));
      SetVisibleWIV(FieldName);

      if iLibre <= 3 then
      begin
        { Valeurs libres }
        FieldName := 'WIV_VALLIBRE' + intToStr(iLibre);
        SetControlProperty('T' + FieldName, 'CAPTION', RechDom('WVALLIBREWIV', 'T' + FieldName, false));
        SetVisibleWIV(FieldName);

        { Dates libres }
        FieldName := 'WIV_DATELIBRE' + intToStr(iLibre);
        SetControlProperty('T' + FieldName, 'CAPTION', RechDom('WDATELIBREWIV', 'T' + FieldName, false));
        SetVisibleWIV(FieldName);

        { Décisions libres }
        FieldName := 'WIV_BOOLLIBRE' + intToStr(iLibre);
        SetControlProperty(FieldName, 'CAPTION', RechDom('WBOOLLIBREWIV' , 'T' + FieldName, false));
        SetControlProperty(FieldName, 'VISIBLE', Copy(RechDom('WBOOLLIBREWIV' , 'T' + FieldName, false), 1, 2) <> '.-');
      end;
    end;
end;


procedure AGLRTChainageCherche(parms:array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFSaisieList) then ToTof:=TFSaisieList(F).LaTof
  else ToTof:=Nil;
  if (ToTof is TOF_CHAINAGES) then TOF_CHAINAGES(ToTof).RTChainageCherche else exit;
end;

procedure AGLRTMenuRechTree(parms:array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFSaisieList) then ToTof:=TFSaisieList(F).LaTof
  else ToTof:=Nil;
  if (ToTof is TOF_CHAINAGES) then TOF_CHAINAGES(ToTof).RTMenuRechTree else exit;
end;

Initialization
registerclasses ( [ TOF_CHAINAGES ] ) ;
RegisterAglProc( 'RTChainageCherche', TRUE , 0, AGLRTChainageCherche);
RegisterAglProc( 'RTMenuRechTree', TRUE , 0, AGLRTMenuRechTree);
end.


