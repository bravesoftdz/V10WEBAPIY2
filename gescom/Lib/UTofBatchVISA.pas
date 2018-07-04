unit UTofBatchVISA;

interface
uses  uTofAfBaseCodeAffaire,StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, HDimension,UTOM,AGLInit,
      Utob,Messages,HStatus,Ent1, Paramsoc,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      Fiche, HDB, mul, DBGrids, db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
      M3VM, M3FP, HTB97, HQry, EntGC, AGLInitGc, UtilGc;

type
    TOF_GCPieceVISA_Mul = class(TOF_AFBASECODEAFFAIRE)
    private
        MasquerColonnes : Boolean;
				RestoreVisa : boolean;
{$IFDEF EAGLCLIENT}
        FPVCol : integer;
{$ENDIF}
        TWC         : String;
        SetPieceVivante : String;
        StPlus 	: string;
        Tiers       : THEdit;
        VenteAchat  : THEdit;
        FactureBTp,factureNegoce : boolean;
    procedure ControleChamp(Champ, Valeur: String);
        procedure GetObjects;
        procedure SetScreenEvents;
        procedure TiersOnElipsisClick(Sender: TObject);

    public
       // Visa des pièces par lot
        procedure BatchVISA;
        procedure SetVISA;
        procedure SetAllVISA;

        procedure UpdateColumns; // Masquage automatique des colonnes Viseur et Date visa
        {$IFDEF EAGLCLIENT}
        procedure RefreshCrossez;
        procedure OnUpdateNext; override;
        {$ELSE}
        procedure OnGetFieldText(Sender: TField; var Text: string; DisplayText: Boolean);
        {$ENDIF}
        procedure OnArgument(stArgument : String); override;
        procedure OnLoad; override;
        procedure OnUpdate; override;
        procedure NomsChampsAffaire( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ; override ;
    end;

implementation
uses BTPUtil;

// procedure appellée par le bouton BVISA
procedure AGLBatchVISA(Parms : Array of Variant; Nb : Integer);
var F : TForm;
    TOTOF : TOF;
begin
F := TForm(Longint(Parms[0]));
if (F is TFmul) then TOTOF := TFMul(F).LaTOF
                else exit;
if (TOTOF is TOF_GCPieceVISA_Mul) then TOF_GCPieceVISA_Mul(TOTOF).BatchVISA
                                  else exit;
end;

Procedure TOF_GCPieceVISA_Mul.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
BEGIN
  Aff           :=THEdit(GetControl('GP_AFFAIRE'))   ;
  Aff0          :=THEdit(GetControl('AFFAIRE0'));
  Aff1          :=THEdit(GetControl('GP_AFFAIRE1')) ;
  Aff2          :=THEdit(GetControl('GP_AFFAIRE2')) ;
  Aff3          :=THEdit(GetControl('GP_AFFAIRE3')) ;
  Aff4          :=THEdit(GetControl('GP_AVENANT'))  ;

  //Tiers:=THEdit(GetControl('GP_TIERS'))   ;
END ;

// true-ization par lot du champ VISA des pièces
procedure TOF_GCPieceVISA_Mul.BatchVISA;
var i : integer;
    DejaVisa : Boolean;
begin
  with TFMul(Ecran) do
   begin
   if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
      begin
      //MessageAlerte('Veuillez sélectionner les pièces à viser');
      if RestoreVisa then
      begin
      	PGIBox('Veuillez sélectionner les pièces à mettre en attente de visa', Caption);
      end else
      begin
      	PGIBox('Veuillez sélectionner les pièces à viser', Caption);
      end;
      exit;
      end;

   if FListe.AllSelected then
      BEGIN
      //if HShowMessage('0;Confirmation;Voulez-vous viser toutes les pièces ?;Q;YN;Y;N;','','')<>mrYes then exit;
      if restoreVisa then
      begin
      	if PGIAsk('Voulez-vous restaurer le visa de toutes les pièces ?', Caption) <> mrYes then exit;
      end else
      begin
      	if PGIAsk('Voulez-vous viser toutes les pièces ?', Caption) <> mrYes then exit;
      end;
      //if Transactions(SetAllVISA,3) <> oeOK then MessageAlerte('Impossible de viser toutes les pièces');
      if Transactions(SetAllVISA,3) <> oeOK then
      begin
        if restoreVisa then
        begin
          PGIBox('Impossible de restaurer le visa de toutes les pièces', Caption);
        end else
        begin
          PGIBox('Impossible de viser toutes les pièces', Caption);
        end;
      end;
      FListe.AllSelected := false;
      END ELSE
      BEGIN
      //if HShowMessage('0;Confirmation;Viser les pièces sélectionnées?;Q;YN;Y;N;','','')<>mrYes then exit;
      if restoreVisa then
      begin
      	if PGIAsk('Voulez-vous restaurer le visa des pièces sélectionnées ?', Caption) <> mrYes then exit;
      end else
      begin
      	if PGIAsk('Voulez-vous viser les pièces sélectionnées ?', Caption) <> mrYes then exit;
      end;

      DejaVisa := False;
      InitMove(FListe.NbSelected,'');
      for i := 0 to FListe.NbSelected-1 do
         BEGIN
         FListe.GotoLeBookMark(i);
        {$IFDEF EAGLCLIENT}
         Q.TQ.Seek(FListe.Row-1) ;
        {$ELSE}
        {$ENDIF}
				 if RestoreVisa then
         begin
           if Q.FindField('GP_ETATVISA').AsString = 'ATT' then // Ne pas restaurer le visa des pièces en attente
           begin
              DejaVisa := True
           end else
           begin
            if Transactions(SetVISA,0) <> oeOK then PGIBox('Impossible de restaurer le visa de la pièce n° '+Q.FindField('GP_NUMERO').AsString, Caption);
           end;
         end else
         begin
           if Q.FindField('GP_ETATVISA').AsString = 'VIS' then // Ne pas viser les pièces déjà visées
              DejaVisa := True
              else
           //if Transactions(SetVISA,3) <> oeOK then MessageAlerte('Impossible de viser la pièce n° '+Q.FindField('GP_NUMERO').AsString);
            if Transactions(SetVISA,0) <> oeOK then PGIBox('Impossible de viser la pièce n° '+Q.FindField('GP_NUMERO').AsString, Caption);
         end;
         MoveCur(False);
         END;
      FListe.ClearSelected;
      FiniMove;
      //if DejaVisa then HShowMessage('1;Information;Certaines pièces déjà visées ont été ignorées;E;O;O;O;','','');
      if DejaVisa then
      begin
      	if RestoreVisa then PGIInfo('Certaines pièces déjà en attente de visa ont été ignorées', Caption)
        							 else PGIInfo('Certaines pièces déjà visées ont été ignorées', Caption);
      end;
      END;
   ChercheClick;
   end;
end;

// Procédure appelée par Transactions pour mettre à Visé le champ GP_ETATVISA
procedure TOF_GCPieceVISA_Mul.SetVISA;
var Nature,DateP,Souche,Nb,Indice : String;
begin
  if RestoreVisa then
  begin
    with TFMul(Ecran) do
    begin
      Nature := Q.FindField('GP_NATUREPIECEG').AsString;
      if (GetInfoParPiece(Nature,'GPP_ACTIONFINI')='VIS') then
        SetPieceVivante:=',GP_VIVANTE="X" ' else SetPieceVivante:='';
      DateP := USDateTime(Q.FindField('GP_DATEPIECE').AsDateTime);
      Souche := Q.FindField('GP_SOUCHE').AsString;
      Nb := Q.FindField('GP_NUMERO').AsString;
      Indice := Q.FindField('GP_INDICEG').AsString;
    end;
    ExecuteSQL('UPDATE PIECE SET GP_ETATVISA="ATT", '+
                                 'GP_VISEUR="", '+
                                 'GP_DATEVISA="'+USDATETIME (idate1900)+ '"' +
                                 SetPieceVivante+
                                 'WHERE '+
                                 'GP_NATUREPIECEG="'+Nature+'" AND '+
                                 'GP_DATEPIECE="'+DateP+'" AND '+
                                 'GP_SOUCHE="'+Souche+'" AND '+
                                 'GP_NUMERO='+Nb+' AND '+
                                 'GP_INDICEG='+Indice);
    if Nature='TRE' then
    begin
      ExecuteSQL('UPDATE PIECE SET GP_ETATVISA="ATT", '+
                                 'GP_VISEUR="", '+
                                 'GP_DATEVISA="'+USDateTime(1900)+'" '+
                                 SetPieceVivante+
                                 'WHERE '+
                                 'GP_NATUREPIECEG="TEM" AND '+
                                 'GP_SOUCHE="'+Souche+'" AND '+
                                 'GP_NUMERO='+Nb+' AND '+
                                 'GP_INDICEG='+Indice+
                                 ' AND GP_VIVANTE="X"');
    end;
  end else
  begin
    with TFMul(Ecran) do
    begin
      Nature := Q.FindField('GP_NATUREPIECEG').AsString;
      if (GetInfoParPiece(Nature,'GPP_ACTIONFINI')='VIS') then
        SetPieceVivante:=',GP_VIVANTE="-" ' else SetPieceVivante:='';
      DateP := USDateTime(Q.FindField('GP_DATEPIECE').AsDateTime);
      Souche := Q.FindField('GP_SOUCHE').AsString;
      Nb := Q.FindField('GP_NUMERO').AsString;
      Indice := Q.FindField('GP_INDICEG').AsString;
    end;
    ExecuteSQL('UPDATE PIECE SET GP_ETATVISA="VIS", '+
                                 'GP_VISEUR="'+V_PGI.User+'", '+
                                 'GP_DATEVISA="'+USDateTime(NowH)+'" '+
                                 SetPieceVivante+
                                 'WHERE '+
                                 'GP_NATUREPIECEG="'+Nature+'" AND '+
                                 'GP_DATEPIECE="'+DateP+'" AND '+
                                 'GP_SOUCHE="'+Souche+'" AND '+
                                 'GP_NUMERO='+Nb+' AND '+
                                 'GP_INDICEG='+Indice);
    if Nature='TRE' then
    begin
      ExecuteSQL('UPDATE PIECE SET GP_ETATVISA="VIS", '+
                                 'GP_VISEUR="'+V_PGI.User+'", '+
                                 'GP_DATEVISA="'+USDateTime(NowH)+'" '+
                                 SetPieceVivante+
                                 'WHERE '+
                                 'GP_NATUREPIECEG="TEM" AND '+
                                 'GP_SOUCHE="'+Souche+'" AND '+
                                 'GP_NUMERO='+Nb+' AND '+
                                 'GP_INDICEG='+Indice+
                                 ' AND GP_VIVANTE="X"');
    end;
  end;
end;

// Procédure appelée par Transactions pour mettre à Visé tous les champs GP_ETATVISA
// en une seule requête
procedure TOF_GCPieceVISA_Mul.SetAllVISA;
var DateVisa : string;
    Q : TQuery;
begin
  if (GetControl('GP_NATUREPIECEG') is THValComboBox) or (Pos(';',GetControlText('GP_NATUREPIECEG'))=0) then  // un seul type de pièce sélectionnné
  begin
    if (GetControlText('GP_NATUREPIECEG') = '')  then
    begin
    PGIBox('Veuillez sélectionner une seule nature de pièces');
    TFMul(Ecran).bSelectAll.Down := false;
    exit;
  end;
  if RestoreVisa then
  begin
    if (GetInfoParPiece(GetControlText('GP_NATUREPIECEG'),'GPP_ACTIONFINI')='VIS') then
      SetPieceVivante:=',GP_VIVANTE="X" ' else SetPieceVivante:='';
    DateVisa := USDateTime(idate1900);
    ExecuteSQL('UPDATE PIECE SET GP_ETATVISA="ATT", '+
                                    'GP_VISEUR="", '+
                                    'GP_DATEVISA="'+DateVisa+'" '+
                                    SetPieceVivante+
                                    TWC +
                                    ' AND GP_ETATVISA="VIS"');
    if GetControlText('GP_NATUREPIECEG')='TRE' then
    begin
      ExecuteSQL('UPDATE PIECE SET GP_ETATVISA="ATT", '+
                                    'GP_VISEUR="", '+
                                    'GP_DATEVISA="'+DateVisa+'" '+
                                    SetPieceVivante+
                                    ' Where GP_NATUREPIECEG="TEM" AND GP_VIVANTE="-"'+
                                    ' And GP_NUMERO IN'+
                                    ' (Select GP_NUMERO From PIECE Where GP_NATUREPIECEG="TRE"'+
                                    ' And GP_ETATVISA="ATT" And GP_VISEUR=""'+
                                    ' And GP_DATEVISA="'+UsDateTime(idate1900)+'")');
    end;
  end else
  begin
    if (GetInfoParPiece(GetControlText('GP_NATUREPIECEG'),'GPP_ACTIONFINI')='VIS') then
      SetPieceVivante:=',GP_VIVANTE="-" ' else SetPieceVivante:='';
    DateVisa := USDateTime(NowH);
    ExecuteSQL('UPDATE PIECE SET GP_ETATVISA="VIS", '+
                                    'GP_VISEUR="'+V_PGI.User+'", '+
                                    'GP_DATEVISA="'+DateVisa+'" '+
                                    SetPieceVivante+
                                    TWC +
                                    ' AND GP_ETATVISA="ATT"');
    if GetControlText('GP_NATUREPIECEG')='TRE' then
    begin
      ExecuteSQL('UPDATE PIECE SET GP_ETATVISA="VIS", '+
                                    'GP_VISEUR="'+V_PGI.User+'", '+
                                    'GP_DATEVISA="'+USDateTime(NowH)+'" '+
                                    SetPieceVivante+
                                    ' Where GP_NATUREPIECEG="TEM" AND GP_VIVANTE="X"'+
                                    ' And GP_NUMERO IN'+
                                    ' (Select GP_NUMERO From PIECE Where GP_NATUREPIECEG="TRE"'+
                                    ' And GP_ETATVISA="VIS" And GP_VISEUR="'+V_PGI.User+'"'+
                                    ' And GP_DATEVISA="'+DateVisa+'")');
    end;
  end;
  end else
  begin
    Q:=TFMul(Ecran).Q;
    Q.First;
    while Not Q.EOF do
    BEGIN
      if RestoreVisa then
      begin
        if Q.FindField('GP_ETATVISA').AsString <> 'ATT' then // Ne pas restaurer le visa des pièces en attente
        begin
          if Transactions(SetVISA,0) <> oeOK then PGIBox('Impossible de restaurer le visa de la pièce n° '+Q.FindField('GP_NUMERO').AsString, Ecran.Caption);
        end;
      end else
      begin
        if Q.FindField('GP_ETATVISA').AsString <> 'VIS' then // Ne pas viser les pièces déjà visées
        begin
          if Transactions(SetVISA,0) <> oeOK then PGIBox('Impossible de viser la pièce n° '+Q.FindField('GP_NUMERO').AsString, Ecran.Caption);
        end;
      end;
      Q.Next;
    end;
  end;
end;

(*
// procedure appellée par l'evenement OnClick du combp GP_ETATVISA
procedure AGLUpdateColumns(Parms : Array of Variant; Nb : Integer);
var F : TForm;
    TOTOF : TOF;
begin
F := TForm(Longint(Parms[0]));
if (F is TFmul) then TOTOF := TFMul(F).LaTOF
                else exit;
if (TOTOF is TOF_GCPieceVISA_Mul) then TOF_GCPieceVISA_Mul(TOTOF).UpdateColumns
                                  else exit;
end;
*)

// Mise à jour de la visibilité des colonnes Viseur et Date visa selon le cas
procedure TOF_GCPieceVISA_Mul.UpdateColumns;
begin
with TFMul(Ecran) do
   begin
   MasquerColonnes := (THValComboBox(FindComponent('GP_ETATVISA')).Value = 'ATT');
   if RestoreVisa then TToolBarButton97(FindComponent('BOUVRIR')).Enabled := (THValComboBox(FindComponent('GP_ETATVISA')).Value <> 'ATT')
   								else TToolBarButton97(FindComponent('BOUVRIR')).Enabled := (THValComboBox(FindComponent('GP_ETATVISA')).Value <> 'VIS');
//   ChercheClick;
   end;
end;

// Masquer les colonnes par défaut (à l'apparition de la fiche)
procedure TOF_GCPieceVISA_Mul.OnArgument(stArgument : String);
Var CC 			: THValComboBox ;
    Critere	: string;
    Champ		: string;
    Valeur	: string;
    i_ind 	: integer;
begin

  fMulDeTraitement  := true;

inherited;

	 fTableName := 'PIECE';
  FactureBTp := false;
  factureNegoce := false;
  //Chargement des zones ecran dans des zones programme
  GetObjects;

	THValComboBox(GetControl('GP_ETATVISA')).Value := 'ATT';
	StPlus  :=  '';

	restoreVisa := false;

	Critere:=(Trim(ReadTokenSt(StArgument)));

  While (Critere <> '') do
  BEGIN
    i_ind:=pos(':',Critere);
    if i_ind = 0 then i_ind:=pos('=',Critere);
    if i_ind <> 0 then
       begin
       Champ:=copy(Critere,1,i_ind-1);
       Valeur:=Copy (Critere,i_ind+1,length(Critere)-i_ind);
       end
    else
       Champ := Critere;
    Controlechamp(Champ, Valeur);
    Critere:=(Trim(ReadTokenSt(StArgument)));
  END;

  //Gestion des évènement des zones écran
  SetScreenEvents;

  //if (ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) and (GetControlText('GP_VENTEACHAT')='ACH') then SetcontrolTExt('GP_NATUREPIECEG','FF');

  {$IFNDEF BTP}
  if Not(ctxGCAFF in V_PGI.PGIContexte) then  SetControlVisible ('TGP_AFFAIRE1', False);
  {$ENDIF}

  CC:=THValComboBox(GetControl('GP_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;
  //
  CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;

end;

Procedure TOF_GCPieceVISA_Mul.ControleChamp(Champ : String;Valeur : String);
Begin
  if Champ='BTP' then
     begin
    FactureBTp := true;
  end else if Champ='NEGOCE' then
  begin
    FactureNegoce := true;
  end else if Champ='RESTORE' then
  begin
     restoreVisa := true;
     ecran.Caption := 'Restauration du visa';
     ecran.Update;
     // Positionnement du mode de visa a 'VISE'
     THValComboBox(GetControl('GP_ETATVISA')).Value := 'VIS';
     StPlus:='';
     end
  else if Champ ='TRANSFERT' then
     SetControlProperty('GP_NATUREPIECEG','Vide',False)
  else if Champ ='VENTEACHAT' then
     Begin
    VenteAchat.Text := Valeur;
     StPlus:= StPlus + 'AND (GPP_VENTEACHAT="' + Valeur + '")';
    if VenteAchat.text = 'ACH' then
      Tiers.Plus   := 'AND (T_NATUREAUXI="FOU") '
    else
      Tiers.Plus   := 'AND (T_NATUREAUXI="CLI") ';
  end;

  If Champ = 'STATUT' then
     Begin
     if (valeur = 'APP') or (valeur = 'INT') then
     Begin
        if GetParamSocSecur('SO_FACTPROV',False) = False then
          SetControlProperty('GP_NATUREPIECEG', 'Plus', ' AND GPP_NATUREPIECEG IN ("FAC","AVC")')
        else
          SetControlProperty('GP_NATUREPIECEG', 'Plus', ' AND GPP_NATUREPIECEG IN ("FPR","FAC","AVC")');
        SetControlProperty('GP_NATUREPIECEG','Enabled',True);
     End;
    //
    if Valeur = 'APP' then
        Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'W')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'W');
         	SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="W"');
      //
          SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
          SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Appel');
          SetControlText('TGP_AFFAIRE1', 'Appel');
        end
     Else if valeur = 'INT' then
          Begin
      //
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'I')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'I');
         	SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="I"');
      //
          SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
          SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Contrat');
          SetControlText('TGP_AFFAIRE1', 'Contrat');
          end
     Else if valeur = 'AFF' then
          Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'A')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'A');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("A", "")');
          SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
          SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Chantier');
          SetControlText('TGP_AFFAIRE1', 'Chantier');
          end
    else if Valeur = 'GRP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then
        SetControlText('XX_WHERE', ' AND AFF_AFFAIRE0 IN ("W","A")')
      else if assigned(GetControl('AFFAIRE0'))  then
        SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("W","A")');
    end
     Else if valeur = 'PRO' then
          Begin
      //
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'P')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'P');
         	SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="P"');
      //
          SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''offre');
          SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Appel d''offre');
          SetControlText('TGP_AFFAIRE1', 'Appel d''offre');
          end
     else
          Begin
          SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
          SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Affaire');
          SetControlText('TGP_AFFAIRE1', 'Affaire');
          end
     end;

end;

procedure TOF_GCPieceVISA_Mul.OnLoad;
begin
inherited;
  if (GetControl('GP_NATUREPIECEG') is ThValComboBox) and (GeTControltext('GP_NATUREPIECEG')<>'') then SetControlEnabled ('GP_NATUREPIECEG',false);
  if (GetControlText('GP_VENTEACHAT')='ACH') then
  begin
    SetControlProperty('GP_NATUREPIECEG','Plus',AfPlusNatureAchat);
    SetControltext('XX_WHERE',GetControlText('XX_WHERE')+AfPlusNatureAchat(false,false));
  end else if (GetControlText('GP_VENTEACHAT')='VEN') then
  begin
    if FactureBTp then
    begin
      SetControlProperty('GP_NATUREPIECEG','Plus',GetPiecesVenteBTP(1));
      SetControltext('XX_WHERE',GetControlText('XX_WHERE')+GetPiecesVenteBTP(1,false));
    end else if factureNegoce then
    begin
      SetControlProperty('GP_NATUREPIECEG','Plus',GetPiecesVenteBTP(2));
      SetControltext('XX_WHERE',GetControlText('XX_WHERE')+GetPiecesVenteBTP(2,false));
    end else
    begin
      SetControlProperty('GP_NATUREPIECEG','Plus',GetPiecesVenteBTP(3));
      SetControltext('XX_WHERE',GetControlText('XX_WHERE')+GetPiecesVenteBTP(3,false));
    end;
  end;
  UpdateColumns;
  TWC := RecupWhereCritere(TPageControl(TFMul(Ecran).Pages));
end;

// Rétablir le masquage des colonnes à chaque updatage sinon elles réapparaissent
procedure TOF_GCPieceVISA_Mul.OnUpdate;
var i : integer;
begin
inherited;
{$IFDEF EAGLCLIENT}
with TFMul(Ecran).FListe do
  for i := 0 to ColCount-1 do
  begin
     if ((UpperCase(Cells[i, 0]) = 'VISA') or
         (UpperCase(Cells[i, 0]) = 'VISEUR') or
         (UpperCase(Cells[i, 0]) = 'DATE VISA'))
        and MasquerColonnes then ColWidths[i] := 0;
     if (UpperCase(Cells[i, 0]) = 'PIèCE VISéE') then FPVCol := i;
  end;
RefreshCrossez;
{$ELSE}
with TFMul(Ecran).FListe do
  for i := 0 to Columns.Count-1 do
    with Columns[i] do
      begin
      if (UpperCase(Title.Caption) = 'VISA') or
         (UpperCase(Title.Caption) = 'VISEUR') or
         (UpperCase(Title.Caption) = 'DATE VISA') then Visible := not MasquerColonnes;
      //if (UpperCase(Title.Caption) = 'VISA') then
         //Field.OnGetText := OnGetFieldText;
      end;
{$ENDIF}
end;

{$IFDEF EAGLCLIENT}
procedure TOF_GCPieceVISA_Mul.OnUpdateNext;
begin
inherited;
RefreshCrossez;
end;

procedure TOF_GCPieceVISA_Mul.RefreshCrossez;
var i, j : integer;
begin
with TFMul(Ecran).FListe do
  for i := 0 to ColCount-1 do
   if (UpperCase(Cells[i, 0]) = 'VISA') then
    for j := 1 to RowCount-1 do
     if Cells[i, j] = '-' then
       if Cells[FPVCol, j] = 'VIS' then Cells[i, j] := 'X'
                                   else Cells[i, j] := ' ';
end;
{$ELSE}
// Interception de l'affichage de la colonne Visa pour afficher les coches
procedure TOF_GCPieceVISA_Mul.OnGetFieldText(Sender: TField; var Text: string; DisplayText: Boolean);
begin
if TFMul(Ecran).Q.FindField('GP_ETATVISA').AsString = 'VIS' then Text := 'X' else Text := '-';
end;
{$ENDIF}



procedure TOF_GCPieceVISA_Mul.GetObjects;
begin

  Tiers         := THEdit(Getcontrol('GP_TIERS'));

  VenteAchat    := THEdit(Getcontrol('GP_VENTEACHAT'));

end;

procedure TOF_GCPieceVISA_Mul.SetScreenEvents;
begin

  Tiers.OnElipsisClick := TiersOnElipsisClick;
  //

end;

procedure TOF_GCPieceVISA_Mul.TiersOnElipsisClick(Sender : TObject);
begin

  if VenteAchat.Text = 'ACH' then
    DispatchRecherche(Tiers,2,'T_NATUREAUXI = "FOU"','','')
  else
    DispatchRecherche(Tiers,2,'T_NATUREAUXI = "CLI" or T_NATUREAUXI = "PRO"','','');

end;


initialization
RegisterClasses([TOF_GCPieceVISA_Mul]);
RegisterAGLProc('BatchVISA',true,0,AGLBatchVISA);   // true = passer la form, 0 = pas d'arg
//RegisterAGLProc('UpdateColumns',true,0,AGLUpdateColumns);
end.
