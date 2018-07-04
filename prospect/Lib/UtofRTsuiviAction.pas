unit UtofRTsuiviAction;

interface
uses  Classes,forms,sysutils,
      HEnt1,UTOF, UTobView,UtilGc ,StdCtrls
      ,Windows    //GetkeyState
{$IFNDEF EAGLCLIENT}
     ,db
     {$IFNDEF DBXPRESS} ,dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
     ,mul
     ,Fe_Main
{$else}
     ,eMul
     ,MainEagl
{$ENDIF}
{$ifdef AFFAIRE}
      ,UtofAfTraducChampLibre
{$ENDIF}
      ,ParamSoc,UtilSelection,UtilRT,utofAfBaseCodeAffaire,hctrls,EntGC,UtilPGI,HMsgBox,utob, stat;

Function RTLanceFiche_RTSuiviAction(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
{$ifdef AFFAIRE}
                //mcd 11/05/2006 12940  pour faire affectation depuis ressource si paramétré
     TOF_RTSuiviAction = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
    TOF_RTSuiviAction = Class (TOF_AFBASECODEAFFAIRE)
{$endif}
{$IFDEF AFFAIRE}
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override;
        procedure bSelectAff1Click(Sender: TObject);     override ;
{$ENDIF AFFAIRE}
     private
         TobViewer1: TTobViewer;
         procedure ActTVOnDblClickCell(Sender: TObject ) ;
         procedure GereLesActionsCompl( Sender : TObject );
         function EstMultiSocEnreg(Sender: TObject) : boolean;
     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad ; override ;
     END ;




implementation
{CRM_MNG_012FQ10807_070408 }
uses wcommuns ;
const

  TexteMessage: array[1..2] of string = (
    {1}'L''action modifiée a été supprimée par un autre utilisateur',
    {2}'Consultation impossible : les données ne se trouvent pas dans la base en cours'
    );

Function RTLanceFiche_RTSuiviAction(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
  AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;


procedure TOF_RTSuiviAction.OnArgument(Arguments : String ) ;
var F : TForm;
    Memo : THMemo;
begin
inherited ;
  if Arguments <> 'GRF' then
    begin
    F := TForm (Ecran);
    MulCreerPagesCL(F,'NOMFIC=GCTIERS');

    if GetParamSocSecur('SO_RTGESTINFOS001',False) = True then
        MulCreerPagesCL(F,'NOMFIC=RTACTIONS');
    end;
  TobViewer1:=TTobViewer(getcontrol('TV'));
  {if (GetControl('MULTIDOSSIER') = nil) then }TobViewer1.OnDblClick:= ActTVOnDblClickCell ;
  if (Ecran.name = 'RTRTSuiviActionMTIE') then
  begin
  setcontroltext   ('MULTIDOSSIER',MS_CODEREGROUPEMENT);
  end;
{$IFDEF AFFAIRE}
     if ( not (ctxAffaire in V_PGI.PGIContexte) ) and ( not ( ctxGCAFF in V_PGI.PGIContexte) ) then
{$ENDIF}
      begin
      SetControlVisible ('BEFFACEAFF1',false); SetControlVisible ('BSELECTAFF1',false);
      SetControlVisible ('TRAC_AFFAIRE',false); SetControlVisible ('RAC_AFFAIRE1',false);
      SetControlVisible ('RAC_AFFAIRE2',false); SetControlVisible ('RAC_AFFAIRE3',false);
      SetControlVisible ('RAC_AVENANT',false);
      end;
memo := THMemo(GetControl('FSQL'));
if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
  if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
  else begin
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
// TJA + PL le 14/08/08 : on fait les choses proprement : on insert au lieu d'écraser sauvagement la ligne 12 !!!
//    Memo.lines[12]:= Memo.lines[12] + ',YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3,YTC_TABLELIBRETIERS1,YTC_TABLELIBRETIERS2,YTC_TABLELIBRETIERS3';
    memo.Lines.Insert(12, ',YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3,YTC_TABLELIBRETIERS1,YTC_TABLELIBRETIERS2,YTC_TABLELIBRETIERS3');
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
 if (GetControl('RAC_OPERATION') <> nil) and Not GetParamsocSecur('SO_AFRTOPERATIONS',False)
    then  begin
    SetControlVisible('RAC_OPERATION',false);
    SetControlVisible('TRAC_OPERATION',false);
    end;
 if (GetControl('RAC_PROJET') <> nil) and Not GetParamsocSecur('SO_RTPROJGESTION',False)
    then  begin
    SetControlVisible('RAC_PROJET',false);
    SetControlVisible('TRAC_PROJET',false);
    end;
 If (Not VH_GC.GaSeria) or not (GetParamSocSecur ('SO_AFRTPROPOS',False)) then
   begin
   if (GetControl('TRAC_AFFAIRE') <> nil) then  SetControlVisible('TRAC_AFFAIRE',false);
   if (GetControl('RAC_AFFAIRE1') <> nil) then  SetControlVisible('RAC_AFFAIRE1',false);
   if (GetControl('RAC_AFFAIRE2') <> nil) then  SetControlVisible('RAC_AFFAIRE2',false);
   if (GetControl('RAC_AFFAIRE3') <> nil) then  SetControlVisible('RAC_AFFAIRE3',false);
   if (GetControl('RAC_AVENANT') <> nil) then  SetControlVisible('RAC_AVENANT',false);
   if (GetControl('BEFFACEAFF1') <> nil) then  SetControlVisible('BEFFACEAFF1',false);
   if (GetControl('BSELECTAFF1') <> nil) then  SetControlVisible('BSELECTAFF1',false);
   end;
 SetControlProperty ('T_NATUREAUXI', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
// TJA + PL le 14/08/08 : on fait les choses proprement : on n'écrase plus la ligne 11, les champs existent dans la fiche, c'est maintenant inutile de les ajouter en dur
// Memo.lines[11]:=',T_PRESCRIPTEUR,T_SOCIETEGROUPE';
{$endif}
end;

procedure TOF_RTSuiviAction.OnLoad;
var Confid : string;
{CRM_MNG_012FQ10807_070408 }
    wheremulti : string;
begin
inherited;
//  if (V_PGI.MenuCourant = 92 ) then Confid:='CON' else Confid:='CONF';
  Confid:='CON';
  SetControlText('XX_WHERE',RTXXWhereConfident(Confid)) ;
{CRM_MNG_012FQ10807_070408 }
  wheremulti:=MulWhereMultiChoix (TForm (Ecran),'RPR',iif( Assigned(TRadioButton(GetControl('MULTIET'))) AND
    (TRadioButton(GetControl('MULTIET')).checked), 'AND','OR'));
  if Assigned(GetControl('XX_WHEREMULTI')) then
    SetControlText('XX_WHEREMULTI',wheremulti);

  if GetParamSocSecur('SO_RTGESTINFOS001',False) then
    begin
    wheremulti:=MulWhereMultiChoix (TForm (Ecran),'RD1',iif( Assigned(TRadioButton(GetControl('MULTIET'))) AND
      (TRadioButton(GetControl('MULTIET')).checked), 'AND','OR'));
    if Assigned(GetControl('XX_WHEREMULTI')) then
      SetControlText('XX_WHEREMULTI',GetControltext('XX_WHEREMULTI')+
        iif(GetControltext('XX_WHEREMULTI')<>'',' AND ','')+wheremulti);
    end;

{fin CRM_MNG_012FQ10807_070408 }
end;

procedure TOF_RTSuiviAction.ActTVOnDblClickCell(Sender: TObject );
var Stchaine : string;
  KeyState              : Short;
  ActionF               : TActionFiche;

begin
with TTobViewer(sender) do
    begin
    keyState            := GetKeyState(VK_CONTROL);          //retourne une valeur <0 si touche [CTRL] enfoncée
    if KeyState < 0 then
      ActionF           := taModif
    else
      ActionF           := taConsult;

    if  (EstMultiSocEnreg (sender) = False) then
        begin
    if (ColName[CurrentCol] = 'RAC_AUXILIAIRE') or (ColName[CurrentCol] = 'T_LIBELLE') or (ColName[CurrentCol] = 'RAC_TIERS')then
      //Ouverture de la fiche TIERS
      //V_PGI.DispatchTT (28,ActionF ,AsString[ColIndex('RAC_AUXILIAIRE'), CurrentRow], '','')  //FQ 10732 //TJA 11/09/2008
      V_PGI.DispatchTT(8, ActionF, AsString[ColIndex('RAC_TIERS'), CurrentRow], '', '')         //FQ 10732 //TJA 11/09/2008
    else if (ColName[CurrentCol] = 'RAC_INTERVENANT') or (ColName[CurrentCol] = 'ARS_LIBELLE') then
//      V_PGI.DispatchTT (9,taConsult ,AsString[ColIndex('RAC_INTERVENANT'), CurrentRow], '','')
    //Ouverture de la fiche ASSISTANT
      V_PGI.DispatchTT (6,taConsult ,AsString[ColIndex('RAC_INTERVENANT'), CurrentRow], '','')
    else if (ColName[CurrentCol] = 'RAC_PROJET') then
        begin
        StChaine := AsString[ColIndex('RAC_PROJET'), CurrentRow];
        if Stchaine <> '' then
          //Ouverture de la fiche PROJET
           V_PGI.DispatchTT (30,taConsult ,AsString[ColIndex('RAC_PROJET'), CurrentRow], '','')
        end
    else if (ColName[CurrentCol] = 'RAC_OPERATION') then
        begin
        StChaine := AsString[ColIndex('RAC_OPERATION'), CurrentRow];
        if Stchaine <> '' then
          //Ouverture de la fiche OPERATION
           V_PGI.DispatchTT (23,taConsult ,AsString[ColIndex('RAC_OPERATION'), CurrentRow], '','')
        end
    else if copy(ColName[CurrentCol], 1, 16) = 'RAC_RACBOOLLIBRE' then
        // PL le 14/08/08 : FQ CRM 10893
        // prise de controle des booléens libres d'actions
         GereLesActionsCompl(Sender)
    else if copy(ColName[CurrentCol], 1, 3) = 'RD1' then
        //prise de controle des complements d'actions
         GereLesActionsCompl(Sender)
    else if copy(ColName[CurrentCol], 1, 2) = 'C_' then
        //Ouverture de la fiche CONTACT
         AGLLanceFiche('YY','YYCONTACT','T;'+AsString[ColIndex('RAC_AUXILIAIRE'), CurrentRow]+';'+IntToStr(AsInteger[ColIndex('RAC_NUMACTION'), CurrentRow]),'','ACTION=CONSULTATION')
    else
        begin
        if copy(ColName[CurrentCol],1,3) = 'RCH' then
          //Ouverture de la fiche ACTIONCHAINEE
           V_PGI.DispatchTT (24,taConsult ,IntToStr(AsInteger[ColIndex('RCH_NUMERO'), CurrentRow]), '','')
        else
          //Ouverture de la fiche ACTION
           V_PGI.DispatchTT (22,ActionF ,AsString[ColIndex('RAC_AUXILIAIRE'), CurrentRow]+';'+IntToStr(AsInteger[ColIndex('RAC_NUMACTION'), CurrentRow]), '',';PRODUITPGI='+AsString[ColIndex('RAC_PRODUITPGI'), CurrentRow]);
            end;
        end;
    end;
end;

{$IFDEF AFFAIRE}
procedure TOF_RTSuiviAction.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('RAC_AFFAIRE'));
Aff1:=THEdit(GetControl('RAC_AFFAIRE1'));
Aff2:=THEdit(GetControl('RAC_AFFAIRE2'));
Aff3:=THEdit(GetControl('RAC_AFFAIRE3'));
Aff4:=THEdit(GetControl('RAC_AVENANT'));
Tiers:=THEdit(GetControl('RAC_TIERS'));
end;

procedure TOF_RTSuiviAction.bSelectAff1Click(Sender: TObject);
begin
    SelectionAffaire (EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4, VH_GC.GASeria , false, '', false, true, true)
end;
{$ENDIF AFFAIRE}

function TOF_RTSuiviAction.EstMultiSocEnreg(Sender: TObject) : boolean;
begin
  with TTobViewer(sender) do
  begin
    Result := false;
    if Assigned (GetControl ('MULTIDOSSIER')) then
    begin
      if THValComboBox(GetControl('MULTIDOSSIER')).Value <> '' then
      begin
        if AsString[ColIndex('SYSDOSSIER'), CurrentRow] <> V_PGI.SchemaName then
        begin
          {2'Consultation impossible : les données ne se trouvent pas dans la base en cours'}
          PgiBox (TexteMessage[2]);
          Result := True;
        end;
      end;
    end;
  end;
end;


procedure TOF_RTSuiviAction.GereLesActionsCompl(Sender: TObject);
var
   LeChamp,LeChamp2 : String;
   sAuxiliaire, sNumAction : string;
   LaValeur : string;
   LaCle : String;
   sReq : string;
   LaLigne, LaSel : integer;
   Lacol : integer;
   TobRd1 : TOB;
   Qtob : Tquery;
//   F : TFStat;
   i : Integer;
   bReponse : boolean;
   TOBInfoCompl : TOB;

begin

//   F := TFstat(Ecran);
   LaLigne := TobViewer1.CurrentRow;
   LaSel := TobViewer1.GetRowSelect;
   With TTobViewer(Sender) do
   begin
      // PL le 14/08/08 : en même temps que FQ CRM 10893...gestion un peu plus propre que précédemment
      if copy(ColName[CurrentCol], 1, 14) = 'RD1_RD1LIBBOOL' then       //les booleens
      begin
        try
          LaCle := AsString[ColIndex('RAC_AUXILIAIRE'), LaLigne]+';'+IntToStr(AsInteger[ColIndex('RAC_NUMACTION'), CurrentRow]);
          LeChamp := ColName[CurrentCol];
          bReponse := AsBoolean [CurrentCol, LaLigne];
          LaValeur := 'X';
          if bReponse then
            LaValeur := '-';

          sReq := 'UPDATE RTINFOS001 SET ' + LeChamp
                  + '=(CASE ' + LeChamp + ' WHEN "X" THEN "-" ELSE "X" END), RD1_UTILISATEUR="' + V_PGI.User +'", RD1_DATEMODIF="' + UsDateTime(now) +'" WHERE RD1_CLEDATA="' + LaCle + '"';

          // on tente l'update au cas où la ligne d'info compl existerait déjà
          if ExecuteSQL(sReq)<>1 then
            // L'update n'a pas fonctionné sans planter, c'est certainement que l'action existe
            // mais pas encore la ligne d'info supp associée => on la créé
            begin
            TOBInfoCompl := TOB.Create('RTINFOS001', nil, -1);
              try
                TOBInfoCompl.InitValeurs(False);
                TOBInfoCompl.PutValue('RD1_CLEDATA', LaCle);
                TOBInfoCompl.PutValue(LeChamp, LaValeur);

                if Not TOBInfoCompl.InsertDB(nil) then
                  V_PGI.IoError := oeUnknown;
              finally
              TOBInfoCompl.Free;
              end;
            end;
          SetValueCell(CurrentCol, LaLigne, LaValeur);      //mise à jour de la cellule pour éviter le rechargement du tobviewer

        except
          {1'L''action modifiée a été supprimée par un autre utilisateur'}
          PGIInfo(TexteMessage[1], Ecran.Caption);
          TFstat(Ecran).BChercheClick(sender);
        end;
      end
      else
      // PL le 14/08/08 : FQ CRM 10893 : on gère le double click sur les booléens libres action
      if copy(ColName[CurrentCol], 1, 16) = 'RAC_RACBOOLLIBRE' then       //les booleens libres des actions
      begin
        try
          LeChamp := ColName[CurrentCol];
          bReponse := AsBoolean [CurrentCol, LaLigne];
          sAuxiliaire := AsString[ColIndex('RAC_AUXILIAIRE'), LaLigne];
          sNumAction := inttostr(AsInteger[ColIndex('RAC_NUMACTION'), LaLigne]);
          LaValeur := 'X';
          if bReponse then
            LaValeur := '-';

          sReq := 'UPDATE ACTIONS SET ' + LeChamp
                  + '=(CASE ' + LeChamp + ' WHEN "X" THEN "-" ELSE "X" END), RAC_UTILISATEUR="' + V_PGI.User +'", RAC_DATEMODIF="' + UsDateTime(now) +'" WHERE RAC_AUXILIAIRE="' + sAuxiliaire + '" AND RAC_NUMACTION=' + sNumAction;

          // on tente l'update au cas où la ligne d'info compl existerait déjà
          if ExecuteSQL(sReq)<>1 then
                V_PGI.IoError := oeUnknown;

          SetValueCell(CurrentCol, LaLigne, LaValeur);      //mise à jour de la cellule pour éviter le rechargement du tobviewer

        except
          {1'L''action modifiée a été supprimée par un autre utilisateur'}
          PGIInfo(TexteMessage[1], Ecran.Caption);
          TFstat(Ecran).BChercheClick(sender);
        end;
      end
      else                                                           // on affiche la fiche des autres compléments
      begin
         LaCle := AsString[ColIndex('RAC_AUXILIAIRE'), LaLigne]+'|'+IntToStr(AsInteger[ColIndex('RAC_NUMACTION'), LaLigne])+
                  ';RAC_TYPEACTION='+AsString[ColIndex('RAC_TYPEACTION'), LaLigne];

         AglLancefiche('RT','RTPARAMCL','','','ACTION=MODIFICATION;FICHEPARAM=RTACTIONS;FICHEINFOS='+Lacle+';FOCUS='+ColName[currentCol]) ;

         TobRd1 := Tob.Create('RTINFOS001', nil, -1);
         LaCle := AsString[ColIndex('RAC_AUXILIAIRE'), LaLigne]+';'+IntToStr(AsInteger[ColIndex('RAC_NUMACTION'), LaLigne]);
         Qtob := OpenSQL('SELECT * FROM RTINFOS001 WHERE RD1_CLEDATA="'+LaCle+'"', True);
         if not TobRd1.SelectDB('',Qtob,False) then      //recherche des données ou création nouvelle fiche
            TobRd1.PutValue('RD1_CLEDATA', LaCle);
         Ferme(Qtob);


         For i := 1 to TobRd1.NbChamps do
         begin
            LeChamp := TobRd1.GetNomChamp(i);

            if  (LeChamp <> 'RD1_BLOCNOTE') and (LeChamp <> 'RD1_CLEDATA') and
                (LeChamp <> 'RD1_DATECREATION') and (LeChamp <> 'RD1_DATEMODIF') and
                (LeChamp <> 'RD1_CREATEUR') and (LeChamp <> 'RD1_UTILISATEUR')   then
            begin
               Lacol := ColIndex(Lechamp);
               LeChamp2 := ColName[LaCol];
               SetValueCell(lacol, LaLigne, TobRd1.GetValue(LeChamp2));
            end;

         end;
         FreeAndNil(TobRd1);


      end;
   end;

//   LaLigne := TobViewer1.GetRowSelect;
//   F.BChercheClick(Ecran);
   TobViewer1.SetRowSelect(LaSel);
end;


Initialization
registerclasses([TOF_RTSuiviAction]);

end.
