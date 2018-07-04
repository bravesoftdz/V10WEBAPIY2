unit UtofRTTIERS_RECH;

interface

uses  Controls,Classes,forms,sysutils,
      HCtrls,HEnt1,HMsgBox,UTOF,UtilSelection,UtilRT, ent1,UtilGC,EntGC,uTob,AglInit,hTB97,
{$IFDEF EAGLCLIENT}
      eMul,
{$ELSE}
      Mul,hDb,
{$ENDIF}
{$ifdef AFFAIRE}
      UtofAfTraducChampLibre,
{$ENDIF}
      EntRT,
      ParamSoc,HQry,HStatus,M3FP;

type 
{$ifdef AFFAIRE}
                //mcd 24/11/2005 pour faire affectation depuis ressource si paramétré
     TOF_RTTIERS_RECH = Class (TOF_AFTRADUCCHAMPLIBRE)
{$else}
     TOF_RTTIERS_RECH = Class (TOF)
{$endif}
        procedure OnArgument(Arguments : String) ; override ;
        procedure OnLoad ; override ;
        function RameneListeCodesT: string;
     Private
        F : TForm;
		    MultiSelect : boolean;
		    procedure BOuvrirClick(Sender: TObject);
		    procedure FLISTE_OnDblCLick(Sender: TObject);
{$ifdef GIGI}
        procedure GINatureExit(Sender: TObject);
{$endif}
     END ;

implementation

Uses
  TiersUtil;

procedure TOF_RTTIERS_RECH.BOuvrirClick(Sender: TObject);
var
  iInd : integer;
  TobRetour,TobRF : TOB;
begin
  with TFMul(Ecran) do
  begin
    if not MultiSelect then
    begin
      FLISTE_OnDblCLick(Sender);
    end else
    begin
      TobRetour := TOB.Create('',nil,-1);
      if not FListe.AllSelected then
      begin
        if FListe.NbSelected = 0 then
        begin
          TobRF := TOB.Create('',TobRetour,-1);
          TobRF.AddChampSupValeur('T_TIERS',Q.FindField('T_TIERS').AsString);
        end else
        begin
          for iInd := 0 to FListe.NbSelected -1 do
          begin
            FListe.GotoLeBookMark(iInd);
    {$IFDEF EAGLCLIENT}
            Q.TQ.Seek (FListe.Row-1) ;
    {$ENDIF}
            TobRF := TOB.Create('',TobRetour,-1);
            TobRF.AddChampSupValeur('T_TIERS',Q.FindField('T_TIERS').AsString);
          end;
        end;
      end else
      begin
        Q.First;
        While not Q.Eof do
        begin
          TobRF := TOB.Create('',TobRetour,-1);
          TobRF.AddChampSupValeur('T_TIERS',Q.FindField('T_TIERS').AsString);
          Q.Next;
        end;
      end;
      TheTob := TobRetour;
      if TheTob.Detail.Count > 0 then
        Retour := TheTob.Detail[0].GetString('T_TIERS');
      Close;
    end;
  end;
end;

procedure TOF_RTTIERS_RECH.FLISTE_OnDblCLick(Sender: TObject);
begin
  if ecran.name='RTCTI_RECHTIERS' then
    TFMul(Ecran).Retour:= GetField('T_TIERS')+';;'+GetField('T_NATUREAUXI')+';'+GetField('T_AUXILIAIRE')
	else
		TFMul(Ecran).Retour := GetField('T_TIERS');

  TFMul(Ecran).Close;
end;

procedure TOF_RTTIERS_RECH.OnArgument(Arguments : String) ;
var xx_where,stArgument, tmp, tmp1, tmp2: string;
    var x : integer;
    st : string ;
    i ,ii, jj: integer;
Begin
{$ifndef GIGI}   //mcd 11/07/06 fait en fin de fct
	inherited ;
{$endif}

  { Gestion de la multisélection ds la liste }
  if pos('MULTISELECTION', Arguments)>0 then
	begin
		Multiselect	:= True;
	  Arguments		:= ReadTokenSt(Arguments);
	  SetControlVisible('BSELECTALL', MultiSelect);
		{$IFNDEF EAGLCLIENT}
    	TFMul(Ecran).FListe.MultiSelection := MultiSelect;
  	{$ELSE}
    	TFMul(Ecran).Fliste.MultiSelect := MultiSelect;
  	{$ENDIF}
	end;

	F := TForm (Ecran);
	stArgument := Arguments;
	stArgument := TransformeLesInToOr(stArgument);
	SetControlProperty('T_TIERS','Tag',-9970) ;
(*if (GetControl('T_NATUREAUXI') <> nil)  then
 begin
 SetControlVisible ('T_NATUREAUXI',False);
 SetControlVisible ('TT_NATUREAUXI',False);
 end;   *)
{$Ifdef GIGI}
	if (stArgument = '') then
  begin
  	stArgument := VH_GC.AfTypTiersGRCGI;
  end
	else
  begin    //mcd 15/03/2005 revu .. il ne faut pas perdre les param autres que T_NatureAuxi
  	i :=  pos('T_NATUREAUXI="AUD"',stArgument);
  	if i <> 0 then
    begin
    	stArgument := Copy(Stargument,0,i-1)+ Copy(StArgument,I+21,StrLen(Pchar(StArgument))); // pour supprimer appel pour saisie apportuer avec AUD et DIV non géré
    end;
  end;
{$endif}
	while (pos(' =', stArgument) > 0) do // pour que le x=pos('T_NATUREAUXI=', stArg) marche bien
	begin
  	System.Delete (stArgument, Pos (' =', stArgument), 1);
	end;
	while (pos('= ', stArgument) > 0) do // pour que le x=pos('T_NATUREAUXI=', stArg) marche bien
	begin
  	System.Delete (stArgument, Pos ('= ', stArgument) + 1, 1);
	end;
	if (stArgument = '') then
  	stArgument := '(T_NATUREAUXI="PRO" or T_NATUREAUXI="CLI")';
	setcontroltext('XX_WHERE',stArgument) ;
	{$IFDEF NOMADE} //Ajout des propspects dans le mul
		if (pos('"PRO"',GetControlText('XX_WHERE')) = 0 ) then SetControlText('XX_WHERE',GetControlText('XX_WHERE')+'OR T_NATUREAUXI="PRO"');
	{$ENDIF}
	if (GetControl('T_NATUREAUXI') <> nil)  then
  begin //mcd 10/02/2005 passe par condition plus de tnatauxi au lieu du xx_where
    // pour la restriction sur les natures.. par contre, on doit garder
    // mes éventuelles autres condition passées dans cette zone (pe t_Ipayeur)
  	xx_where :=  GetControlText('XX_WHERE');
  	Tmp :='';
    //on regarde si ( dans la formule .. dans ce cas, on considère qu'il
    //il y a des info en plus de T_natureauxi
    //si pas de (), on considère qu'il n'y a que des info sur T_NatureAuxi
    //les cas répertorié de where
        // t_natureauxi="CLI"
        // (t_natureauxi="cli" or t_natureauxi_"pro")
        // t_ispayeur="X" and (t_natureauxi="cli" or t_natureauxi_"pro")
    //11/05/2005 nouveau cas
        // t_ispayeur="X" and t_natureAuxi="FOU"
    //donc les () ne sont pas obligatoire.. il faut traiter ce nouveau cas
  	i := Pos ('(',xx_where);
  	if i <> 0 then
    begin //on cherche la ) fermante
    	ii :=Pos (')',xx_where);
      //on regarde si entre les 2 ( ), cela concerne le T_NATUREAUXI
    	jj :=  Pos('T_NATUREAUXI',Copy (xx_where,i,ii));
    	if JJ<>0 then
      begin
       	//on garde la condition autre que NatureAuxi
      	tmp :=Copy (xx_where,0,i-1) ;
      	if (ii+1) < int(Strlen(Pchar(xx_where))) then tmp :=tmp+' ' + Copy(xx_where,ii+1,Strlen(Pchar(xx_where)));
      	xx_where := Copy(XX_Where,i,ii);// on ne garde que la nature auxi
      end;
    end
  	else
    begin  //mcd 11/05/2005
    	i := Pos ('T_NATUREAUXI',xx_where);
    	if i > 1 then
      	xx_where :=  Copy(XX_Where,i,Strlen(Pchar(xx_where)));
    end;
      //pour TTNatAuxi, on remplace NatureAuxi par code
  	xx_where :=  StringReplace(xx_where, 'T_NATUREAUXI', 'CO_CODE', [rfReplaceAll]);
    // BDU - 11/04/07 - FQ : 13921. Ajout de ( et ) autour de xx_where
    // car si il y a plusieurs T_NATUREAUXI on se retrouve avec une requête
    // CO_TYPE = "NTT" AND CO_CODE = "XXX" OR CO_CODE = "YYY" OR CO_CODE = "ZZZ"
  	SetControlProperty ('T_NATUREAUXI', 'Plus', ' AND (' +xx_where + ')');
	  (* mcd 13/05/05 mécanisme changé, car si on a plusieurs natures passées dans l'argument
  	   ne les met pas en critère dans T_Natureauxi
    	 SetControlText ('T_NATUREAUXI','CLI');
	     if (pos('FOU',xx_where) <>0) and (pos('CLI',xx_where) =0) then  SetControlText ('T_NATUREAUXI','FOU');
  	*)
  	i:=0  ;
  	tmp1:='';
  	i :=pos ('=',copy(xx_where,i,strlen(Pchar(xx_where)))) ;
  	tmp2:=xx_where;
  	While  (i >0)  //on récupere tous les type de client mis dans la requête
    do
    begin
      tmp1:= tmp1+Copy(tmp2,i+2,3) + ';';
      jj := strlen(Pchar(tmp2));
      tmp2:=Copy(tmp2,i+2,jj-i);
      i := pos ('=',tmp2) ;
    end;
  	SetControlText ('T_NATUREAUXI',tmp1);

		//mcd 07/04/2006 : si jamais il y a un esapce dans tmps, l'AGL ajout AND dans la clause where ..
  	SetControlText('XX_WHERE',trim(tmp)) ;  //on ne garde que les info # de t_natureauxi
  end;
	{pas de fournisseurs}
	if (pos('T_NATUREAUXI="FOU"',stArgument) = 0 ) then
  begin
		{$ifdef GIGI}
  	for ii := 1 to 3 do  //si pas FOU , pas les champs dans la vue associé à la fiche
    	SetControlText('YTC_TABLELIBREFOU' + IntToStr(ii),'');
		{$endif}
  	if (GetControlText('XX_WHERE') = '') then
    	SetControlText('XX_WHERE',RTXXWhereConfident('CON',true))
  	else
    begin
    	xx_where := GetControlText('XX_WHERE');
    	xx_where := xx_where + RTXXWhereConfident('CON',true);
    	SetControlText('XX_WHERE',xx_where) ;
    end;
  end
	else
	{ fournisseurs sans clients }
  if (pos('T_NATUREAUXI="CLI"',stArgument) = 0 ) then
  begin
    if assigned(GetControl('T_PARTICULIER')) then
      SetControlVisible('T_PARTICULIER', False);
  	TFMul(F).SetDBListe('RFMULTIERSRECH');
		{$ifdef AFFAIRE}  //mcd 10/01/07 12914 change GIGI en AFFAIRE
  		for ii := 1 to 9 do  //si sur FOU uniquement, pas les champs dans la vue associé à la fiche
    		SetControlText('YTC_TABLELIBRETIERS' + IntToStr(ii),'');
  		SetControlText('YTC_TABLELIBRETIERSA','');
		{$endif}
  	if (GetControlText('XX_WHERE') = '') then
    	SetControlText('XX_WHERE',RTXXWhereConfident('CONF',true))
  	else
    begin
    	xx_where := GetControlText('XX_WHERE');
    	xx_where := xx_where + RTXXWhereConfident('CONF',true);
    	SetControlText('XX_WHERE',xx_where) ;
    end;
  end
  else
  	{ fournisseurs et clients : par exemple duplication tiers }
	//    TFMul(F).Q.Liste:='RTMULTITIERSRECH';
  begin
    TFMul(F).SetDBListe('RTMULTITIERSRECH');
      //mcd 29/05/2006 ajout confidentialité pour que ce soit OK
      //car les info ne sont plus passée depuis la fiche appelante.. pb sinon sur ce mul de rech sur TTnatTiers
    tmp:= RTXXWhereConfident('CON',true) ;
    tmp2 := RTXXWhereConfident('CONF',true);
    if (Tmp <>'') or (tmp2 <>'') then
     begin
     xx_where:='';
     xx_where := xx_where+'(((T_NATUREAUXI="CLI") OR (T_NATUREAUXI="PRO"))  ' + tmp+')';
     xx_where := xx_where+' OR ';
     xx_where := xx_where+'(T_NATUREAUXI="FOU" ' + tmp2+')';
     SetControlText('XX_WHERE',xx_where) ;
     end;
  end;

	x:=pos('T_NATUREAUXI=',stArgument) ;
	st:='';
	if x<>0 then st:=copy(stArgument,x+14,3);
	if VH_GC.GCIfDefCEGID then
  	if (st='CLI') and (pos('T_NATUREAUXI="PRO"',stArgument) <> 0 ) then st:='PRO' ; // si cli et PRO on force NATUREAUXI à PRO pour creation de prospect par defaut
	setcontroltext('NATUREAUXI',st) ;
	if (ctxScot in V_PGI.PGIContexte) and (st ='FOU') then
  begin
  	Ecran.Caption := Traduirememoire('Recherche d''un fournisseur');
   	UpdateCaption(Ecran);
    SetControlText ('TT_TIERS','Fournisseur');//mcd 11/07/06
  end;

{	if V_PGI.MenuCourant = 30 then // JTR - eQualité 11591
	begin
  	if (pos('"CLI"',stArgument) > 0) and (Not AutoriseCreationTiers ('CLI')) then
  	begin
    	SetControlVisible('BINSERT',False) ;
    	SetControlVisible('B_DUPLICATION',False) ;
  	end;
	end else
	begin
  	if (st='CLI' ) and (Not AutoriseCreationTiers ('CLI')) and
       (Not AutoriseCreationTiers ('PRO')) then
  	begin
    	SetControlVisible('BINSERT',False) ;
    	SetControlVisible('B_DUPLICATION',False) ;
  	end;
  	if (st='PRO' ) and (Not AutoriseCreationTiers ('CLI')) and
  		 (Not AutoriseCreationTiers ('PRO')) then
    	SetControlVisible('BINSERT',False) ;

  	if (st='FOU' ) and (Not AutoriseCreationTiers ('FOU')) then
  	begin
    	SetControlVisible('BINSERT',False) ;
    	SetControlVisible('B_DUPLICATION',False) ;
  	end;
  end;      }
  if (st <> 'FOU' ) then
	begin
    if (ctxGRC in V_PGI.PGIContexte) then
      begin
      if ((Not AutoriseCreationTiers ('CLI')) and
         (Not AutoriseCreationTiers ('PRO'))) or
         (GetParamSocSecur('SO_RTCONFIDENTIALITE',False) = True and VH_RT.RTExisteConfident = False) then SetControlVisible('BINSERT',False) ;
      end
    else
      begin
      if Not AutoriseCreationTiers ('CLI') then SetControlVisible('BINSERT',False) ;
      end;
	end
  else
  begin
  	if (Not AutoriseCreationTiers ('FOU')) or (GetParamSocSecur('SO_RFCONFIDENTIALITE',False) = True and VH_RT.RFExisteConfident = False) then
    	SetControlVisible('BINSERT',False) ;
  end;

	{$ifdef GIGI}
	if (pos('T_NATUREAUXI="CLI"',stArgument) <> 0 ) or (pos('T_NATUREAUXI="PRO"',stArgument) <> 0 ) or (pos('T_NATUREAUXI="NCP"',stArgument) <> 0 ) then
	{$else}
	if (pos('T_NATUREAUXI="CLI"',stArgument) <> 0 ) or (pos('T_NATUREAUXI="PRO"',stArgument) <> 0 ) then
	{$endif}
  begin
  	SetControlVisible ('PCOMPLEMENT', True);
    //fait autrement 03/2005  SetControlProperty('T_TIERS','PLUS','(T_NATUREAUXI="CLI" OR T_NATUREAUXI="PRO")'); //JS F10292
  	MulCreerPagesCL(F,'NOMFIC=GCTIERS');
  	if (pos('T_NATUREAUXI="FOU"',stArgument) = 0 ) then
    begin
    	for i:=1 to 3 do
      	SetControlProperty('YTC_TABLELIBREFOU'+IntToStr(i),'Tag',-9970) ;
    end;
  end;
	if (pos('T_NATUREAUXI="FOU"',stArgument) <> 0 ) then
  begin
  	SetControlVisible ('PCOMPLEMENTFOUR', True);
  	SetControlProperty('T_TIERS','PLUS','T_NATUREAUXI="FOU"'); //JS F10292
  	GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'YTC_TABLELIBREFOU', 3, '');
  	if (pos('T_NATUREAUXI="CLI"',stArgument) = 0 ) then
    begin
    	SetControlVisible ('TT_ENSEIGNE', False);
    	SetControlVisible ('T_ENSEIGNE', False);
    	SetControlVisible ('TT_REPRESENTANT', False);
    	SetControlVisible ('T_REPRESENTANT', False);
    	SetControlVisible ('TT_SOCIETEGROUPE', False);
    	SetControlVisible ('T_SOCIETEGROUPE', False);
    	SetControlVisible ('TT_PRESCRIPTEUR', False);
    	SetControlVisible ('T_PRESCRIPTEUR', False);
    	SetControlVisible ('T_ZONECOM', False);
    	SetControlVisible ('TT_ZONECOM', False);
    	//SetControlVisible ('LRAISONSOCIALE', False);
    	//SetControlVisible ('RAISONSOCIALE', False);
    	SetControlProperty('T_ENSEIGNE','Tag',-9970) ;
    	SetControlProperty('T_REPRESENTANT','Tag',-9970) ;
    	SetControlProperty('T_SOCIETEGROUPE','Tag',-9970) ;
    	SetControlProperty('T_PRESCRIPTEUR','Tag',-9970) ;
    	SetControlProperty('T_ZONECOM','Tag',-9970) ;
    	for i:=1 to 9 do
      	SetControlProperty('YTC_TABLELIBRETIERS'+IntToStr(i),'Tag',-9970) ;
    	SetControlProperty('YTC_TABLELIBRETIERSA','Tag',-9970) ;
    end;
    if GetParamSocSecur('SO_RTGESTINFOS003',False) = True then
      MulCreerPagesCL(F,'NOMFIC=GCFOURNISSEURS');
  end;
	if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
  	if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
  	else
    begin
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
    ThMultiValCOmboBox(Getcontrol('T_NatureAuxi')).onexit := GiNatureExit;
	{$endif}

  { Evènements }
  if Assigned(GetControl('BOUVRIR')) then
    TToolBarButton97(GetControl('BOUVRIR')).OnClick := BOuvrirClick;
  if Assigned(GetControl('FLISTE')) then
    THGrid(GetControl('FLISTE')).OnDblClick := FLISTE_OnDblCLick;
{$ifdef GIGI}
inherited;      //mcd 11/07/06 déplacer pour OK affectation automatique valeur mul si fournisseur t_natureauxi doit déjà être affecté)
{$endif}
End ;

procedure TOF_RTTIERS_RECH.OnLoad;
begin
{$IFDEF CTI}
  if (copy (F.Name,1,15) = 'RTCTI_RECHTIERS') and (GetControlText ('APPELANT') = '') then
     SetControlText('APPELANT',GetControlText('TELEPHONE'));
{$ENDIF}
end;

function AGLRameneListeCodesT(parms: array of variant; nb: integer): variant;
var
  F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then
    TOTOF := TFMul(F).LaTOF
  else
    exit;
  if (TOTOF is TOF_RTTIERS_RECH) then
    result := TOF_RTTIERS_RECH(TOTOF).RameneListeCodesT
  else
    exit;
end;

function TOF_RTTIERS_RECH.RameneListeCodesT: string;
var
  F: TFMul;
  Q: THQuery;
  i: integer;
{$IFDEF EAGLCLIENT}
  L: THGrid;
{$ELSE}
  L: THDBGrid;
{$ENDIF}
  code: string;
begin
  Result := '';
  F := TFMul(Ecran);
  L := F.FListe;
  Q := F.Q;

{$IFDEF EAGLCLIENT}
  if F.bSelectAll.Down then
    if not F.Fetchlestous then
      begin
      F.bSelectAllClick(nil);
      F.bSelectAll.Down := False;
      exit;
      end else
        F.Fliste.AllSelected := true;
{$ENDIF}

  if L.AllSelected then
    begin
    Q.First;
    while not Q.Eof do
    begin
      code:=Q.FindField('T_TIERS').asstring ;
      if Result = '' then Result:=code else Result:=Result+';'+code;
      Q.Next;
    end;
    L.AllSelected := False;
    end
  else
    if F.FListe.NbSelected <> 0 then
      begin
      InitMove(L.NbSelected, '');
      for i := 0 to L.NbSelected - 1 do
      begin
        MoveCur(False);
        L.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
        Q.TQ.Seek(L.Row - 1);
{$ENDIF}
        code:=TFmul(Ecran).Q.FindField('T_TIERS').asstring ;
        if Result = '' then Result:=code else Result:=Result+';'+code;
      end;
      L.ClearSelected;
      end;
  FiniMove;
end;

{$ifdef GIGI}
procedure TOF_RTTIERS_RECH.GINatureExit(Sender: TObject);
var st : string;
begin
    //mcd 17/01/2007 13628 si on change la valuer de t_NatureAuxi, il faut aussi la changer
    //dans NtaureAuxi, afin que la création se fasse sur le bonne valeur
  st:=copy(GetCOntrolText('T_natureAuxi'),1,3);
	setcontroltext('NATUREAUXI',st) ;
end;
{$endif}


Initialization
registerclasses([TOF_RTTIERS_RECH]);
RegisterAglFunc('RameneListeCodesT', TRUE, 0, AGLRameneListeCodesT);
end.
