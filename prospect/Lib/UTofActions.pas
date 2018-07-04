unit UTofActions;

interface
uses  windows,Classes,forms,sysutils,
      HCtrls,HEnt1,UTOF,ParamSoc,UtilRT,TiersUtil,aglInit,EntRT,Hqry, StdCtrls,
{$IFDEF EAGLCLIENT}
      MaineAGL,eMul,UTob,
{$ELSE}
      HDB,db,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}Fe_Main,Mul,
{$ENDIF}
      HMsgBox,HStatus,M3Fp,utofAfBaseCodeAffaire,EntGC,HTB97,UtilAction,UtilSelection;
Type
     TOF_Actions = Class (TOF_AFBASECODEAFFAIRE)
{$IFDEF AFFAIRE}
     procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override;
     procedure bSelectAff1Click(Sender: TObject);     override ;
{$ENDIF AFFAIRE}
     private
        stProduitPgi : string;
{$IFDEF EAGLCLIENT}
    		Fliste 			: THGrid;
{$ELSE}
    		Fliste 			: THDbGrid;
{$ENDIF}
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        procedure BPiece_OnClick(Sender: TObject);
		    procedure GSDblClick(Sender: Tobject);
        procedure BinsertClick (Sender : Tobject);
    procedure RAC_TIERS_OnChange(Sender: Tobject);
    procedure BDUPLICATION_OnClick(Sender: Tobject);
        procedure ControleChamp(Champ, Valeur: String);

     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad ; override ;
        procedure OnClose ; override ;
        procedure RTMajDateRappelAction(Rappel : String);
     END ;

var VerrouModif :boolean;

Procedure RTLanceFiche_Actions_Mul (Nat,Cod : String ; Range,Lequel,Argument : string) ;

implementation

Procedure RTLanceFiche_Actions_Mul(Nat,Cod : String ; Range,Lequel,Argument : string) ;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_Actions.OnArgument(Arguments : String ) ;
var x       :integer;
  StrArgs : TStringList;
    i       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin
	fMulDeTraitement := true;
inherited ;

{$IFDEF EAGLCLIENT}
	Fliste:= THGrid (GetCOntrol('Fliste'));
{$ELSE}
	Fliste:= THDbGrid (GetCOntrol('Fliste'));
{$ENDIF}


  //FV1 : 05/12/20014 - C'est quoi cette merde de gestion des paramètres !!!!!
  //Remise en place de la version standard

  //TJA 28/06/2007      pour gérer les arguments
  {sTmp := StringReplace(Arguments, ';', chr(VK_RETURN), [rfReplaceAll]);
  // On réceptionne les paramètres passés par l'écran appelant, s'il y en a
  StrArgs := TStringList.Create;
  StrArgs.Text := STmp;
  if (StrArgs.Count = 0) then
  begin
    StrArgs.Free;
    StrArgs := nil;
  end;}
  //FV1 : 05/12/20014
  Repeat
    Critere:=uppercase(ReadTokenSt(Arguments)) ;
    valeur := '';
    if Critere<>'' then
    begin
      x:=pos('=',Critere);
      if x<>0 then
      begin
        Champ:=copy(Critere,1,x-1);
        Valeur:=copy(Critere,x+1,length(Critere));
      end
      else
        Champ := Critere;
      ControleChamp(Champ, Valeur);
  end;
  until  Critere='';

  stProduitPgi := '';

  if copy(TFMul(Ecran).name,1,2) = 'RT' then
   begin
    stProduitPgi:='GRC';
    for i:=1 to 3 do
        SetControlCaption('TRAC_TABLELIBRE'+IntToStr(i),RechDom('RTLIBCHAMPSLIBRES','AL'+IntToStr(i),FALSE)) ;
        if GetParamsocSecur('SO_RTACTGESTECH',False) = FALSE then
         begin
         SetControlEnabled('RAC_DATEECHEANCE',FALSE) ;
         SetControlEnabled('RAC_DATEECHEANCE_',FALSE) ;
         SetControlEnabled('TRAC_DATEECHEANCE',FALSE) ;
         SetControlEnabled('TRAC_DATEECHEANCE_',FALSE) ;
         end;
  End
  else if copy(TFMul(Ecran).name,1,2) = 'RF' then
   begin
    stProduitPgi:='GRF';
        if GetParamsocSecur('SO_RFACTGESTECH',False) = FALSE then
         begin
         SetControlEnabled('RAC_DATEECHEANCE',FALSE) ;
         SetControlEnabled('RAC_DATEECHEANCE_',FALSE) ;
         SetControlEnabled('TRAC_DATEECHEANCE',FALSE) ;
         SetControlEnabled('TRAC_DATEECHEANCE_',FALSE) ;
         end;
   end;

  //if (V_PGI.MenuCourant = 92 ) and (TFMul(Ecran).name <> 'RTACTIONS_ACTGEN') then
  //FV1 : 05/12/2014 - Tout est en double c'est quoi ce bordel !!!!
  {if (ecran <> Nil) and (TFMul(Ecran).name = 'RTACTIONS_TIERS') then
  begin
  TFMul(Ecran).OnKeyDown:=FormKeyDown ;
  if Assigned (GetControl ('VOIRFICHE')) then
    begin  //mcd 24/01/2006 10392
    if GetSynRegKey('RTVoirActionTiers','-',TRUE) = 'X' then SetControlChecked ('VOIRFICHE',true)
    else SetControlChecked ('VOIRFICHE',False);
    if Not GetParamSocSecur ('SO_RTACTCLIENT',false) then SetControlVisible ('VOIRFICHE',false);
    end;
  {mng 18/10/07 FQ 10741
  VerrouModif := False;
  if not VH_RT.RTCreatActions then
    begin
    x := pos('CONSULTATION',Arguments);
    if (x <> 0) then VerrouModif :=true;
    end;
  end;}
  //FV1 : 05/12/2014 -

  if (ecran <> Nil) and (TFMul(Ecran).name = 'RFACTIONS_TIERS') then
  begin
  TFMul(Ecran).OnKeyDown:=FormKeyDown ;
  if Assigned (GetControl ('VOIRFICHE')) then
    begin  //mcd 24/01/2006 10392
    if GetSynRegKey('RFVoirActionFOUR','-',TRUE) = 'X' then SetControlChecked ('VOIRFICHE',true)
    else SetControlChecked ('VOIRFICHE',False);
    if Not GetParamSocSecur ('SO_RFACTFOURNISSEUR',false) then SetControlVisible ('VOIRFICHE',false);
    end;
    end;


  if (ecran <> Nil) and ((TFMul(Ecran).name= 'RTACTIONS_MUL_RAP')  or
                        (TFMul(Ecran).name = 'RTACTIONS_MUL')      or
                        (TFMul(Ecran).name = 'RTACTIONS_DUPLI')    or
                        (TFMul(Ecran).name = 'RFACTIONS_MUL')      or
                        (TFMul(Ecran).name = 'RFACTIONS_DUPLI')) then
    begin
    	// Modif BTP
			Fliste.OnDblClick := GSDblClick;
    if assigned (TToolbarButton97(GetControl('Binsert')))       then TToolbarButton97(GetControl('Binsert')).OnClick := BinsertClick;
    if assigned (ThEdit(GetCOntrol('RAC_TIERS')))               then ThEdit(GetCOntrol('RAC_TIERS')).OnChange :=RAC_TIERS_OnChange;
    if assigned (TToolbarButton97(GetControl('BDUPLICATION')))  then TToolbarButton97(GetControl('BDUPLICATION')).OnClick := BDUPLICATION_OnClick;
    if assigned (ThEdit(GetControl('RAC_INTERVENANT')))         then SetControlText ('RAC_INTERVENANT',VH_RT.RTResponsable);
      end;

  {$IFDEF AFFAIRE}
  if ( not (ctxAffaire in V_PGI.PGIContexte) ) and ( not ( ctxGCAFF in V_PGI.PGIContexte) ) then
  {$ENDIF}
      begin
      SetControlVisible ('BEFFACEAFF1',false); SetControlVisible ('BSELECTAFF1',false);
      SetControlVisible ('TRAC_AFFAIRE',false); SetControlVisible ('RAC_AFFAIRE1',false);
      SetControlVisible ('RAC_AFFAIRE2',false); SetControlVisible ('RAC_AFFAIRE3',false);
      SetControlVisible ('RAC_AVENANT',false);
      end
   else
      begin
    {$IFDEF AFFAIRE}
      if pos('NOCHANGEAFFAIRE',Arguments) <> 0 then
        begin
        SetControlEnabled ('BEFFACEAFF1',false); SetControlEnabled ('BSELECTAFF1',false);
        SetControlEnabled ('TRAC_AFFAIRE0',false); SetControlEnabled ('RAC_AFFAIRE1',false);
        SetControlEnabled ('RAC_AFFAIRE2',false); SetControlEnabled ('RAC_AFFAIRE3',false);
        SetControlEnabled ('RAC_AVENANT',false);
        end
    {$ENDIF}
   end;

  if Assigned(GetControl('BPIECE')) then TToolBarButton97(GetControl('BPIECE')).OnClick := BPiece_OnClick;

  {$Ifdef GIGI}
  if (GetControl('RAC_OPERATION') <> nil) and Not GetParamsocSecur('SO_AFRTOPERATIONS',False) then
  begin
    SetControlVisible('RAC_OPERATION',false);
    SetControlVisible('TRAC_OPERATION',false);
    end;

  if (GetControl('RAC_PROJET') <> nil) and Not GetParamsocSecur('SO_RTPROJGESTION',False) then
  begin
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
  {$endif}

  {$IFDEF GRCLIGHT}
  if ( copy(TFMul(Ecran).name,1,2) = 'RT' ) and (not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False)) then
    begin
    SetControlVisible('RAC_OPERATION',false);
    SetControlVisible('TRAC_OPERATION',false);
    SetControlVisible('RAC_PROJET',false);
    SetControlVisible('TRAC_PROJET',false);
    end;
  {$ENDIF GRCLIGHT}

  if (ecran <> Nil) and ((TFMul(Ecran).name = 'RTACTIONS_MUL') or (TFMul(Ecran).name = 'RTACTIONS_TIERS')) then MulCreerPagesCL(Ecran,'NOMFIC=RTACTIONS');

 //TJA 28/06/2007
 if (StrArgs <> nil) then
    // Si une StringList de réception a été passée en paramètre
    // la variable de réception de l'écran Ressource est remplie
  if (StrArgs.Values['MONTIERS'] <> '') then SetControlText('MONTIERS', StrArgs.Values['MONTIERS']);

  StrArgs.Free;

end;

procedure TOF_Actions.ControleChamp(Champ, Valeur: String);
begin

  VerrouModif := False;
  if not VH_RT.RFCreatActions then
  begin
    if champ = 'CONSULTATION' then VerrouModif :=true;
  end;

  //mcd 21/12/2005 equalite 10392
  if champ = 'PREVU'          then SetControlText ('RAC_ETATACTION','PRE');

  If champ = 'STATUT' then
  begin
    if valeur = 'APP' then
    begin
      if assigned(GetControl('CHANTIER0')) then SetControlText('CHANTIER0', 'W');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Appel');
      SetControlText('TRAC_AFFAIRE', 'Code Appel');
    end
    Else if valeur = 'INT' then
    begin
      if assigned(GetControl('CHANTIER0')) then SetControlText('CHANTIER0', 'I');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Contrat');
      SetControlText('TRAC_AFFAIRE', 'Code Contrat');
    end
    Else if valeur = 'AFF' then
    begin
      if assigned(GetControl('CHANTIER0')) then SetControlText('CHANTIER0', 'A');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Chantier');
      SetControlText('TRAC_AFFAIRE', 'Code chantier');
    end
    Else if valeur = 'PRO' then
    begin
      if assigned(GetControl('CHANTIER0')) then SetControlText('CHANTIER0', 'P');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''Offre');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Appel d''offre');
      SetControlText('TRAC_AFFAIRE', 'Code Appel d''Offre');
    end
    Else
    Begin
      if assigned(GetControl('CHANTIER0')) then SetControlText('CHANTIER0', '');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Affaire');
      SetControlText('TRAC_AFFAIRE', 'Code affaire');
    end;
  end;

end;

procedure TOF_Actions.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var    Q : TQuery ;
       stAuxi,stNatureauxi,stLibelle,stParticulier : string;
       TypeAction : TActionFiche ;
begin
Case Key of
    VK_F12 : {encours} if ((ssAlt in Shift) and (ssShift in Shift)) then begin; Key:=0 ; AfficheRisqueClientDetail(getControlText('RAC_TIERS')); end
                       else if (ssAlt in Shift) then begin Key:=0 ; AfficheRisqueClient(getControlText('RAC_TIERS')); end;
    VK_F10 : {Contacts} if (ssAlt in Shift) and (getControlText('RAC_TIERS')<>'') then
                        begin
                        if VerrouModif then TypeAction:=taConsult else TypeAction:=taModif;
                        Key:=0 ;
                        stAuxi:=TiersAuxiliaire (getControlText('RAC_TIERS'));
                        Q := OpenSQL('SELECT T_NATUREAUXI,T_PARTICULIER,T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="'+stAuxi+'"', True);
                        if Q.Eof then begin Ferme(Q); exit; end
                          else begin
                          stNatureauxi:=Q.FindField('T_NATUREAUXI').asstring;
                          stLibelle:=Q.FindField('T_NATUREAUXI').asstring;
                          stParticulier:=Q.FindField('T_NATUREAUXI').asstring;
                          end;
                        Ferme(Q);
                        AGLLanceFiche('YY','YYCONTACT','T;'+TiersAuxiliaire (getControlText('RAC_TIERS')),'', ActionToString(TypeAction)+';TYPE=T;'+'TYPE2='+stNatureauxi+';PART='+stParticulier+';TITRE='+stLibelle);
                        end;

END;
if (ecran <> nil) then TFMul(ecran).FormKeyDown(Sender,Key,Shift);
end;

procedure TOF_Actions.OnLoad;
var xx_where,Confid : string ;
begin
  inherited;

  if (TFMul(Ecran).name <> 'RTACTIONS_MUL_ASS') and (TFMul(Ecran).name <> 'RTACTIONS_DUPLI_T') and
     (TFMul(Ecran).name <> 'RTACTIONS_MUL_PRO') and (TFMul(Ecran).name <> 'RTACTIONS_TIERS') and
     (TFMul(Ecran).name <> 'RTACT_MUL_CONTACT') and (TFMul(Ecran).name <> 'RFACTIONS_TIERS') and
     (TFMul(Ecran).name <> 'RTACTIONS_AFFAIRE') and (TFMul(Ecran).name <> 'RFACT_MUL_CONTACT') and
     (TFMul(Ecran).name <> 'RFACTIONS_DUPLI_T') then
  begin
    if stProduitpgi <> '' then
    begin
  //    if (V_PGI.MenuCourant = 92 ) or ( ( copy(TFMul(Ecran).name,1,17) = 'RTACTIONS_MUL_RAP') ) then Confid:='CON' else Confid:='CONF';
    if (stProduitpgi = 'GRC') then Confid:='CON' else Confid:='CONF';
    if (GetControlText('XX_WHERE') = '') then
        SetControlText('XX_WHERE',RTXXWhereConfident(Confid))
    else
        begin
        xx_where := GetControlText('XX_WHERE');
        xx_where := xx_where + RTXXWhereConfident(Confid);
        SetControlText('XX_WHERE',xx_where) ;
        end;
    end;
  end;

  if Assigned(GetControl('XX_JOIN')) then
    begin
    if (GetControlText('RAI_RESSOURCE') <> '') then
        SetControlText('XX_JOIN','LEFT JOIN ACTIONINTERVENANT on rac_auxiliaire=rai_auxiliaire and rac_numaction=rai_numaction ')
      else
        SetControlText('XX_JOIN','');
    end;

  if (TFMul(Ecran).name = 'RTACTIONS_TIERS') or (TFMul(Ecran).name = 'RFACTIONS_TIERS') then
  begin
    SetControlChecked('BVERROU',VerrouModif);
    SetControlEnabled('BINSERT',not VerrouModif);
    SetControlEnabled('BDUPLICATION',not VerrouModif);
  end;
end;

procedure TOF_Actions.OnClose;
//mcd 24/01/2006 10392
begin
if (ecran <> Nil) and (TFMul(Ecran).name = 'RTACTIONS_TIERS') then
  begin
  if Assigned (GetControl ('VOIRFICHE')) then
    begin  //mcd 24/01/2006 10392
    If TcheckBox(GetControl('VOIRFICHE')).checked then SaveSynRegKey('RTVoirActionTiers','X',TRUE)
    else SaveSynRegKey('RTVoirActionTiers','-',TRUE);
    end;
  end;
if (ecran <> Nil) and (TFMul(Ecran).name = 'RFACTIONS_TIERS') then
  begin
  if Assigned (GetControl ('VOIRFICHE')) then
    begin  //mcd 24/01/2006 10392
    If TcheckBox(GetControl('VOIRFICHE')).checked then SaveSynRegKey('RFVoirActionFour','X',TRUE)
    else SaveSynRegKey('RFVoirActionFour','-',TRUE);
    end;
  end;

end;

Procedure TOF_Actions.RTMajDateRappelAction( Rappel : String);
var HeureAct : TDatetime;
    Sql : String;
    F : TFMul ;
    i : integer;
{$IFDEF EAGLCLIENT}
       L : THGrid;
{$ELSE}
       L : THDBGrid;
{$ENDIF}
    Q : THQuery;
begin
//  F:=TFMul(Ecran);
  F:=TFMul(Ecran) ;
  if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
     begin
     PGIInfo('Aucun élément sélectionné','');
     exit;
     end;

{$IFDEF EAGLCLIENT}
  if F.bSelectAll.Down then
     if not F.Fetchlestous then
       begin
       F.bSelectAllClick(Nil);
       F.bSelectAll.Down := False;
       exit;
       end else
       F.Fliste.AllSelected := true;
{$ENDIF}
  { calcul de l'heure + délai de répétition }
  if Rappel < '001' then
     HeureAct:=Now+EncodeTime(0, StrToInt(copy(Rappel,2,2)), 0, 0)
  else
    if Rappel < '024' then
      HeureAct:=Now+EncodeTime(StrToInt(Rappel), 0, 0, 0)
    else
      HeureAct:=PlusDate(Now, StrToInt(Rappel) Div 24,'J');

  L:= F.FListe;
  Q:= F.Q;

  if L.AllSelected then
  begin
    InitMove(Q.RecordCount,'');
    Q.First;
    while Not Q.EOF do
    begin
      MoveCur(False);
      Sql:='UPDATE ACTIONS SET RAC_DATERAPPEL="'+FormatDateTime('mm/dd/yyyy hh:nn:ss',HeureAct);
      Sql:=Sql+'" WHERE RAC_AUXILIAIRE="'+TFmul(Ecran).Q.FindField('RAC_AUXILIAIRE').asstring;
      Sql:=Sql+'" AND RAC_NUMACTION='+IntToStr(TFmul(Ecran).Q.FindField('RAC_NUMACTION').asInteger);
      ExecuteSQL(Sql);
      Q.Next;
    end;
    L.AllSelected:=False;
  end else
  begin
    InitMove(L.NbSelected,'');
    for i:=0 to L.NbSelected-1 do
    begin
      MoveCur(False);
      L.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(L.Row-1) ;
{$ENDIF}
      Sql:='UPDATE ACTIONS SET RAC_DATERAPPEL="'+FormatDateTime('mm/dd/yyyy hh:nn:ss',HeureAct);
      Sql:=Sql+'" WHERE RAC_AUXILIAIRE="'+TFmul(Ecran).Q.FindField('RAC_AUXILIAIRE').asstring;
      Sql:=Sql+'" AND RAC_NUMACTION='+IntToStr(TFmul(Ecran).Q.FindField('RAC_NUMACTION').asInteger);
      ExecuteSQL(Sql);
    end;
  L.ClearSelected;
  end;

  if F.bSelectAll.Down then
    F.bSelectAll.Down := False;

  FiniMove;
end;

{$IFDEF AFFAIRE}
procedure TOF_Actions.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('RAC_AFFAIRE'));
Aff1:=THEdit(GetControl('RAC_AFFAIRE1'));
Aff2:=THEdit(GetControl('RAC_AFFAIRE2'));
Aff3:=THEdit(GetControl('RAC_AFFAIRE3'));
Aff4:=THEdit(GetControl('RAC_AVENANT'));
Tiers:=THEdit(GetControl('RAC_TIERS'));
end;

procedure TOF_Actions.bSelectAff1Click(Sender: TObject);
begin
{$IFNDEF BTP}
    SelectionAffaire (EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4, VH_GC.GASeria , false, '', false, true, true)
{$ELSE}
//    SelectionAffaire (EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4)
    SelectionAffaire (EditTiers,EditAff,THEdit(GetControl('CHANTIER0')),EditAff1,EditAff2,EditAff3,EditAff4)
{$ENDIF}

end;
{$ENDIF AFFAIRE}

procedure TOF_Actions.BPiece_OnClick(Sender: TObject);
var stProduitPgi : String;
begin
  { en attendant que vpgi_menu marche dans le cas d'un affichage menu detail }
  if copy(TFMul(Ecran).name,1,2) = 'RT' then
     stProduitPgi:='GRC' else stProduitPgi:='GRF' ;
  if assigned( TFmul(Ecran).Q.FindField('RAC_NUMCHAINAGE') ) then
    begin
{$IFDEF EAGLCLIENT}
    TFmul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;
{$ENDIF}
    if TFmul(Ecran).Q.FindField('RAC_NUMCHAINAGE').asInteger <> 0 then
      RTAffichePieceLiee(TFmul(Ecran).Q.FindField('RAC_NUMCHAINAGE').asInteger,stProduitPgi)
    else
      PGIBox(TraduireMemoire('Il n''y a pas de chaînage pour cette action'),TraduireMemoire('Accès pièce(s) du chaînage'));
    end
  else
    PGIBox(TraduireMemoire('Vous devez mettre le numéro de chaînage dans la liste'),TraduireMemoire('Accès pièce(s) du chaînage'));
end;

// ------------------------------------------------------------------------------------
// --- Recupération des traitements de la fiche (Scripts) pour fonctionnement dnas BTP
// ------------------------------------------------------------------------------------
procedure TOF_Actions.GSDblClick (Sender : Tobject);
Var stBlocProspect : string;
		Auxiliaire,Tiers,TypeAction,Intervenant: string;
    NumAction : integer;
begin
{$IFDEF EAGLCLIENT}
	if FListe.RowCount = 0  then exit;
	TFMul(ecran).Q.TQ.Seek(FListe.Row-1) ;

	Auxiliaire:=TFMul(ecran).Q.FindField('RAC_AUXILIAIRE').AsString;
  Tiers := TFMul(ecran).Q.FindField('RAC_TIERS').AsString;
  TypeAction := TFMul(ecran).Q.FindField('RAC_TYPEACTION').AsString;
  Intervenant := TFMul(ecran).Q.FindField('RAC_INTERVENANT').AsString;
  NumAction := TFMul(ecran).Q.FindField('RAC_NUMACTION').AsInteger;
{$ELSE}
	if FListe.datasource.DataSet.RecordCount = 0  then exit;

	Auxiliaire:=Fliste.datasource.dataset.FindField('RAC_AUXILIAIRE').AsString;
  Tiers := Fliste.datasource.dataset.FindField('RAC_TIERS').AsString;
  TypeAction := Fliste.datasource.dataset.FindField('RAC_TYPEACTION').AsString;
  Intervenant := Fliste.datasource.dataset.FindField('RAC_INTERVENANT').AsString;
  NumAction := Fliste.datasource.dataset.FindField('RAC_NUMACTION').AsInteger;
{$ENDIF}

  //
	if Auxiliaire <> '' then
  begin
    stBlocProspect:='ACTION=MODIFICATION';
    if (RTDroitModifActions(Tiers,TypeAction,Intervenant)=False) then stBlocProspect:= 'ACTION=CONSULTATION';
    if assigned (ThEdit(GetCOntrol('RAC_PROJET'))) then
    begin
    	if ThEdit(GetCOntrol('RAC_PROJET')).text <> '' then stBlocProspect:=stBlocProspect + ';RAC_PROJET='+ThEdit(GetCOntrol('RAC_PROJET')).text;
    end;
    if Assigned (THedit(getCOntrol('RAC_TIERS'))) then
    begin
    	if THedit(getCOntrol('RAC_TIERS')).text <> '' then stBlocProspect:=stBlocProspect + ';RAC_TIERS='+THedit(getCOntrol('RAC_TIERS')).text;
    	if THedit(getCOntrol('RAC_TIERS')).enabled=false then stBlocProspect:=stBlocProspect+';NOCHANGEPROSPECT';
    end;
    if Assigned (THedit(getCOntrol('RAC_PROJET'))) then
    begin
    	if ThEdit(GetCOntrol('RAC_PROJET')).enabled=false then stBlocProspect:=stBlocProspect+';NOCHANGEPROJET';
    end;
    AglLanceFiche('RT','RTACTIONS','',Auxiliaire+';'+IntToStr(NumAction),stBlocProspect) ;
    TtoolBarButton97(GetCOntrol('Bcherche')).Click;
  end;
end;

procedure TOF_Actions.BinsertClick (Sender : Tobject);
var stBlocProspect : string;
begin
    stBlocProspect:='' ;
    if ThEdit(GetCOntrol('RAC_PROJET')).text <> '' then stBlocProspect:=stBlocProspect + ';RAC_PROJET='+ThEdit(GetCOntrol('RAC_PROJET')).text;
    if THedit(getCOntrol('RAC_TIERS')).text <> '' then stBlocProspect:=stBlocProspect + ';RAC_TIERS='+THedit(getCOntrol('RAC_TIERS')).text;
    if THedit(getCOntrol('RAC_TIERS')).enabled=false then stBlocProspect:=stBlocProspect+';NOCHANGEPROSPECT';
    if ThEdit(GetCOntrol('RAC_PROJET')).enabled=false then stBlocProspect:=stBlocProspect+';NOCHANGEPROJET';
    AglLanceFiche('RT','RTACTIONS','','','MONOFICHE;ACTION=CREATION'+stBlocProspect) ;
  TtoolBarButton97(GetCOntrol('Bcherche')).Click;
end;

procedure TOF_Actions.RAC_TIERS_OnChange(Sender : Tobject);
begin
if ( GetParamSocSecur('SO_RTPROJMULTITIERS',true)=false ) then
   THEdit(GetControl('RAC_PROJET')).plus:=TiersAuxiliaire(THEdit(GetControl('RAC_TIERS')).text,False);
end;

procedure TOF_Actions.BDUPLICATION_OnClick( Sender : Tobject);
var Auxiliaire : string;
		NumAction : integer;
begin
{$IFDEF EAGLCLIENT}
	TFMul(ecran).Q.TQ.Seek(TFMul(ecran).FListe.Row-1) ;
	Auxiliaire:=TFMul(ecran).Q.FindField('RAC_AUXILIAIRE').AsString;
  NumAction := TFMul(ecran).Q.FindField('RAC_NUMACTION').AsInteger;
{$ELSE}
	Auxiliaire:=Fliste.datasource.dataset.FindField('RAC_AUXILIAIRE').AsString;
  NumAction := Fliste.datasource.dataset.FindField('RAC_NUMACTION').AsInteger;
{$ENDIF}


// mng 26-05-03 : bouton caché plutot que supprimé
	if (Auxiliaire <> '') then
  begin
    AglLanceFiche('RT','RTACTIONS','','','ACTION=CREATION;DUPLICATION='+Auxiliaire+'|'+InttoStr(NumAction)) ;
  	TtoolBarButton97(GetCOntrol('Bcherche')).Click;
  end;
end;

// ------------------------------------------------------------------------------------
// --- FIN Recupération des traitements de la fiche (Scripts) pour fonctionnement dnas BTP
// ------------------------------------------------------------------------------------

procedure AGLRTMajDateRappelAction(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_Actions) then TOF_Actions(TOTOF).RTMajDateRappelAction(Parms[1]);
end;

Initialization
registerclasses([TOF_Actions]);
RegisterAglProc( 'RTMajDateRappelAction', true , 2, AGLRTMajDateRappelAction);
end.
