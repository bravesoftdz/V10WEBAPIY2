{***********UNITE*************************************************
Auteur  ...... : F. Vautrain
Créé le ...... : 19/10/2005
Modifié le ... :   /  /
Description .. : Facturation des lignes de conso...
Mots clefs ... : TOF;BTFACCONSO
*****************************************************************}
Unit UTofFactConso ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     EntGC,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     DBGrids,
     Hqry,
{$ENDIF}
     HTB97,
     HDB,
     UTOB,
     forms,
     vierge,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,

     AffaireUtil,
     FactGrp,
     ParamSoc,
     FactTiers,
     FactTOB,
     FactUtil,
     FactComm,
     FactPiece,
     FactArticle,
     FactFormule,
     CalcOLEGenericBTP,
     UTOF,uEntCommun ;


Type
  TOF_BTFACCONSO = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private

    //
    //Variables correspondant aux objets de la liste
    BOuvrir1   	 : TToolBarButton97;
    BSELECTAFF1  : TToolBarButton97;
    BEFFACEAFF   : TToolBarButton97;
    BEFFACETIERS : TToolBarButton97;
    BSELECTCON	 : TToolBarButton97;
    BEFFACECON   : TToolBarButton97;
    //
    TFV          : TFVierge ;
    fDateFact    : THEdit;
    //

    BCO_AFFAIRE  : THEdit;
    BCO_AFFAIRE0 : THEdit;
    BCO_AFFAIRE1 : THEdit;
    BCO_AFFAIRE2 : THEdit;
    BCO_AFFAIRE3 : THEdit;
    BCO_AVENANT  : THEdit;
    T_TIERS			 : THEdit;
    BCO_AFFAIRESAISIE : THEdit;
    CONTRAT0 : THEdit;
    CONTRAT1 : THEdit;
    CONTRAT2 : THEdit;
    CONTRAT3 : THEdit;
    CON_AVENANT  : THEdit;

    Affaire0		 : String;

  	//Procedure propre à l'appli
		Procedure ChargementConsoLigne(TobConso, TOBPiece, TobLigne : Tob);
    Procedure ControleChamp(Champ, Valeur: String);
    Procedure ControleCritere(Critere: String);
    Procedure MajConsoLigne(TObResult : Tob);
    Procedure MajEtatAffaire(TobResult: TOB);

    //procedure liée à un objet de la liste
    procedure BEffaceClick(Sender: TOBJect);
    procedure BEffaceTierClick(Sender: TOBJect);
    procedure BSelectAffClick(Sender: TOBJect);
    procedure fDateFactOnExit(Sender: TOBJect);
    procedure GenerationFactConso(Sender: TOBJect);
    procedure ValideFactconso(Sender: TOBJect);
    procedure BEffaceConClick(Sender: TOBJect);
    procedure BSelectConClick(Sender: TOBJect);

  end ;

Implementation
uses BTStructChampSup,
     factdomaines,
     PiecesRecalculs,
     SaisUtil,
     AGLInitBtp,
     UtilTOBPiece,
     UtilPgi;


procedure TOF_BTFACCONSO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTFACCONSO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTFACCONSO.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTFACCONSO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTFACCONSO.OnArgument (S : String ) ;
var Critere : String;
    Valeur  : String;
    Champ   : String;
    X       : integer;
    CC      : THValComboBox;
begin
  Inherited ;
	MemoriseChampsSupLigneETL ('FAC',true);
  MemoriseChampsSupLigneOUV ('FAC');
  MemoriseChampsSupPIECETRAIT;

  //Récupération valeur de argument
	Critere:=(Trim(ReadTokenSt(S)));

  while (Critere <> '') do
    begin
      if Critere <> '' then
      begin
        X := pos (':', Critere) ;
        if x = 0 then
          X := pos ('=', Critere) ;
        if x <> 0 then
        begin
          Champ := copy (Critere, 1, X - 1) ;
          Valeur := Copy (Critere, X + 1, length (Critere) - X) ;
        	ControleChamp(champ, valeur);
				end
      end;
      ControleCritere(Critere);
      Critere := (Trim(ReadTokenSt(S)));
    end;


    // gestion Etablissement (BTP)
    CC:=THValComboBox(GetControl('BTBETABLISSEMENT')) ;
    if CC<>Nil then PositionneEtabUser(CC) ;

    // gestion Domaine (BTP)
    CC:=THValComboBox(GetControl('AFF_DOMAINE')) ;
    if CC<>Nil then PositionneDomaineUser(CC) ;

    BOuvrir1 := TToolbarButton97(ecran.FindComponent('BOuvrir1'));
		BOuvrir1.OnClick := GenerationFactConso;
    BOuvrir1.hint := 'Facturer les lignes de consommations';

    BEffaceAFF := TToolbarButton97(ecran.FindComponent('BEffaceAff'));
    BEffaceAFF.OnClick := BEffaceClick;

    BEffaceTiers := TToolbarButton97(ecran.FindComponent('BEffaceTier'));
    BEffaceTiers.OnClick := BEffaceTierClick;

    // saisie par chantier
	  BSelectAFF1  := TToolBarButton97(ecran.FindComponent('BSELECTAFF1'));
    BSelectAFF1.OnClick := BSelectAffClick;

    //Saisie par contrat
	  BSelectCON  := TToolBarButton97(ecran.FindComponent('BSELECTCON'));
    BSelectCON.OnClick := BSelectConClick;

    BEffaceCON := TToolbarButton97(ecran.FindComponent('BEffaceCon'));
    BEffaceCON.OnClick := BEffaceConClick;

  	BCO_AFFAIRE  := THEdit(GetControl('BCO_AFFAIRE'));
  	BCO_AFFAIRE0 := THEdit(GetControl('BCO_AFFAIRE0'));
  	BCO_AFFAIRE1 := THEdit(GetControl('BCO_AFFAIRE1'));
  	BCO_AFFAIRE2 := THEdit(GetControl('BCO_AFFAIRE2'));
  	BCO_AFFAIRE3 := THEdit(GetControl('BCO_AFFAIRE3'));
  	BCO_AVENANT  := THEdit(GetControl('AFF_AVENANT'));

  	T_TIERS  := THEdit(GetControl('T_TIERS'));

   	BCO_AFFAIRESAISIE  := THEdit(GetControl('BCO_AFFAIRESAISIE'));
  	CONTRAT0 := THEdit(GetControl('CONTRAT0'));
  	CONTRAT1 := THEdit(GetControl('CONTRAT1'));
  	CONTRAT2 := THEdit(GetControl('CONTRAT2'));
  	CONTRAT3 := THEdit(GetControl('CONTRAT3'));
  	CON_AVENANT  := THEdit(GetControl('CON_AVENANT'));

    BCO_AFFAIRE0.text := Affaire0;

   	ChargeCleAffaire(BCO_AFFAIRE0, BCO_AFFAIRE1, BCO_AFFAIRE2, BCO_AFFAIRE3, BCO_AVENANT, BSelectAFF1, TaModif, BCO_AFFAIRE.Text, false);
		//Formatage du code contrat
    ChargeCleAffaire(CONTRAT0, CONTRAT1, CONTRAT2, CONTRAT3, CON_AVENANT, BSelectCON, TaModif, BCO_AFFAIRESAISIE.Text, False);

end ;

Procedure TOF_BTFACCONSO.ControleCritere(Critere : String);
Begin

end;

Procedure TOF_BTFACCONSO.ControleChamp(Champ : String;Valeur : String);
Begin

  if Champ = 'ACTION' then
  Begin
     if Valeur = 'CREATION' Then
     		TFVierge(ecran).TypeAction := TaCreat
     else if Valeur = 'MODIFICATION' then
     		TFVierge(ecran).TypeAction := TaModif
  end;

	if Champ ='STATUT' then
     Begin
  	 if Valeur = 'APP' then Affaire0 := 'W'
  	 Else if Valeur = 'AFF' then Affaire0 := 'A'
     Else if Valeur = 'INT' then Affaire0 := 'I'
     Else if Valeur = 'PRO' then Affaire0 := 'P'
     else Affaire0 := 'A';
     end;
end;

procedure TOF_BTFACCONSO.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTFACCONSO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTFACCONSO.OnCancel () ;
begin
  Inherited ;
end ;

Procedure TOF_BTFACCONSO.BSelectAffClick(Sender : TOBJect);
var Aff0				 : String;
    Aff1				 : String;
    Aff2				 : String;
    Aff3				 : String;
    Aff4				 : String;
    Affaire 		 : string;
    EtatAffaire	 : String;
    bProposition : Boolean;
begin
  Aff0:=affaire0;
  Aff1:='';
  Aff2:='';
  Aff3:='';
  Aff4:='';
  Affaire :='';
  bProposition := False;

  if (BCO_AFFAIRE0 <> Nil) then
     Begin
     if BCO_AFFAIRE0.Text = 'P' then
        bProposition := true
     else if BCO_AFFAIRE0.Text = 'W' then
        EtatAffaire := 'REA';
     end;

  if BCO_AFFAIRE <> Nil then Affaire := BCO_AFFAIRE.Text;

  if not GetAffaireEnteteSt(BCO_AFFAIRE0,BCO_AFFAIRE1,BCO_AFFAIRE2,BCO_AFFAIRE3,BCO_AVENANT,T_TIERS,affaire,bProposition,false,false,true,false,'',false,true,false,EtatAffaire ) then Affaire := '';

  BCO_AFFAIRE.text := Affaire;

  BTPCodeAffaireDecoupe(affaire,Aff0, Aff1,Aff2,Aff3,aff4,taCreat,false);

  BCO_AFFAIRE0.text := Aff0;
  BCO_AFFAIRE1.text := Aff1;
  BCO_AFFAIRE2.text := Aff2;
  BCO_AFFAIRE3.text := Aff3;

end;

Procedure TOF_BTFACCONSO.BEffaceClick(Sender : TOBJect);
Begin
	BCO_AFFAIRE.text  := '';
	BCO_AFFAIRE1.Text := '';
 	BCO_AFFAIRE2.Text := '';
 	BCO_AFFAIRE3.Text := '';
 	BCO_AVENANT.Text  := '';
  T_TIERS.Text := '';
end;

Procedure TOF_BTFACCONSO.BSelectConClick(Sender : TOBJect);
var Aff0				 : String;
    Aff1				 : String;
    Aff2				 : String;
    Aff3				 : String;
    Aff4				 : String;
    Contrat 		 : string;
    Etatcontrat	 : String;
    bProposition : Boolean;
begin

  Aff0:='I';
  Aff1:='';
  Aff2:='';
  Aff3:='';
  Aff4:='';
  contrat :='';

  bProposition := False;

  EtatContrat := 'ENC';

  if BCO_AFFAIRESAISIE <> Nil then Contrat := BCO_AFFAIRESAISIE.Text;

  CONTRAT0.text := Aff0;
  if not GetAffaireEnteteSt(CONTRAT0,CONTRAT1,CONTRAT2,CONTRAT3,CON_AVENANT,T_TIERS,Contrat,bProposition,false,false,true,false,'',false,true,false,EtatContrat ) then Contrat := '';

  BCO_AFFAIRESAISIE.text := Contrat;

  BTPCodeAffaireDecoupe(Contrat,Aff0, Aff1,Aff2,Aff3,aff4,taCreat,false);

  CONTRAT0.text := Aff0;
  CONTRAT1.text := Aff1;
  CONTRAT2.text := Aff2;
  CONTRAT3.text := Aff3;

end;

Procedure TOF_BTFACCONSO.BEffaceConClick(Sender : TOBJect);
Begin
	BCO_AFFAIRESAISIE.text  := '';
	CONTRAT1.Text := '';
 	CONTRAT2.Text := '';
 	CONTRAT3.Text := '';
 	CON_AVENANT.Text  := '';
end;

Procedure TOF_BTFACCONSO.BEffaceTierClick(Sender : TOBJect);
Begin

	BCO_AFFAIRE.text  := '';
	BCO_AFFAIRE1.Text := '';
 	BCO_AFFAIRE2.Text := '';
 	BCO_AFFAIRE3.Text := '';
 	BCO_AVENANT.Text  := '';

	T_TIERS.Text := '';

end;

Procedure TOF_BTFACCONSO.GenerationFactConso(Sender : TOBJect);
Var fCaption      : THLabel;
begin

  //controle si au moins un éléments sélectionné
	if (TFMul(Ecran).FListe.NbSelected=0)and(not TFMul(Ecran).FListe.AllSelected) then
  	 begin
	   PGIInfo('Aucun élément sélectionné','');
  	 exit;
   	 end;

  //Demande de confirmation du traitement
	//if PGIAsk('Confirmez-vous le traitement ?','') <> mrYes then exit ;
  //
  TFV        := TFVierge.create(ecran);
  TFV.Caption:= 'Validation Facuration';
  TFV.height := 150;
  TFV.width  := 310;
  TFV.Top    := round(ecran.height / 2);
  TFV.Left   := round(ecran.width / 2);

  TFV.BValider.OnClick := ValideFactconso;

  fCaption := THLabel.create (TFV);
  fCaption.parent  := TFV;
  fcaption.caption := 'Renseignez la date de facturation';
  fCaption.Autosize:= True;
  fCaption.Top     := 16;
  fCaption.Left    := 60;

  fDateFact := THEdit.create (TFV);
  fDateFact.parent        := TFV;
  fDateFact.ControlerDate := True;
  fDateFact.DefaultDate   := OdDate;
  fDateFact.MaxLength     := 10;
  fDateFact.OpeType       := otDate;
  fDateFact.EditMask      := '!99/99/0000;1;_';
  fDateFact.Top     := 45;
  fDateFact.Left    := 90;
  fdatefact.text    := DateToStr(V_PGI.DateEntree);
  fDateFact.OnExit  := fDateFactOnExit;

  fCaption.visible  := True;
  fDateFact.Visible := True;

  TFV.show;

end;

//FV1 - 20/07/2017 : FS#2635 - VEODIS GROUP - aucun contrôle de date de génération de facture sur un exercice clôturé
Procedure TOF_BTFACCONSO.fDateFactOnExit(Sender : TOBJect);
begin

  //Contrôle que la date de saisie ne soit pas sur un exercice comptable fermé
  if isExerciceClo(StrToDateTime(fdatefact.text)) then
  begin
    PGIInfo(TraduireMemoire('Cette pièce est sur un exercice clôturé en comptabilité.'),'ATTENTION');
    fdateFact.text := DateToStr(V_PGI.DateEntree);
  end;

end;

Procedure TOF_BTFACCONSO.ValideFactconso(Sender : TOBJect);
var	req      : String;
		TobConso : Tob;
    TOBLigne : Tob;
    TOBPiece : TOB;
    TobResult: Tob;
    NumConso : Double;
    DateConso: TDateTime;
    Indice   : Integer;
    F 			 : TFMul;
    i 			 : integer;
    QQ			 : TQuery;
    QConso	 : TQuery;

    //Modif tri des appels par numéro si groupement
    StOrder  : String;
{$IFDEF EAGLCLIENT}
    L 			 : THGrid;
{$ELSE}
    L 			 : THDBGrid;
{$ENDIF}

begin
  Inherited ;

  //Contrôle que la date de saisie ne soit pas sur un exercice comptable fermé
  if isExerciceClo(StrToDateTime(fdatefact.text)) then
  begin
    PGIInfo(TraduireMemoire('Cette pièce est sur un exercice clôturé en comptabilité.'),'ATTENTION');
    fdateFact.text := DateToStr(V_PGI.DateEntree);
    exit;
  end;

  F := TFMul(Ecran);

  TOBConso  := Tob.create('CONSOMMATIONS', nil, -1);
  TOBLigne  := Tob.create('LES LIGNES', nil, -1);
  TOBResult := Tob.create('LES PIECES', nil, -1);
  TOBPiece  := TOB.Create ('LES PIECES',nil,-1);

	L:= TFMul(Ecran).FListe;

  SourisSablier;
  TRY
	if L.AllSelected then
  	 begin
  	 QQ:=F.Q;
     QQ.First;
     while not QQ.EOF do
     	Begin
      TOBConso.InitValeurs;
      //chargement clé
      DateConso:= QQ.findfield('BCO_DATEMOUV').asDateTime;
      NumConso := QQ.findfield('BCO_NUMMOUV').AsFloat;
      Indice   := QQ.findfield('BCO_INDICE').asInteger;
      Req := 'SELECT *, T_AUXILIAIRE, AFF_TIERS FROM CONSOMMATIONS '+
             'LEFT JOIN AFFAIRE ON BCO_AFFAIRE = AFF_AFFAIRE '+
						 'LEFT JOIN TIERS ON AFF_TIERS = T_TIERS '+
						 'WHERE BCO_DATEMOUV="'+USDATETIME(DateConso)+'" AND '+
             'BCO_NUMMOUV='+FloatToStr(NumConso)+' AND '+
             'BCO_INDICE='+IntToStr(Indice);
      Qconso := OpenSQL(Req,true);
      if not QConso.eof then
      begin
        TOBConso.SelectDB ('',Qconso);
   			ChargementConsoLigne(TobConso, TOBPiece,TobLigne);
      end;
      Ferme (QConso);
      //
     	QQ.next;
      end;
     end
  else
     begin
     for i:=0 to L.NbSelected-1 do
         begin
	       L.GotoLeBookmark(i);
         NumConso := TFMul(F).Fliste.datasource.dataset.FindField('BCO_NUMMOUV').AsFloat;
         DateConso:= TFMul(F).Fliste.datasource.dataset.FindField('BCO_DATEMOUV').asDateTime;
    	   Indice   := TFMul(F).Fliste.datasource.dataset.FindField('BCO_INDICE').asInteger;
		     Req := 'SELECT *, T_AUXILIAIRE, AFF_TIERS FROM CONSOMMATIONS '+
                'LEFT JOIN AFFAIRE ON BCO_AFFAIRE = AFF_AFFAIRE '+
					      'LEFT JOIN TIERS ON AFF_TIERS = T_TIERS '+
						    'WHERE BCO_DATEMOUV="'+USDATETIME(DateConso)+'" AND '+
                'BCO_NUMMOUV='+FloatToStr(NumConso)+' AND '+
                'BCO_INDICE='+IntToStr(Indice) + StOrder;
         Qconso := OpenSQL(Req,true);
         if not QConso.eof then
         		begin
            TOBConso.SelectDB ('',Qconso);
   			    ChargementConsoLigne(TobConso, TOBPiece,TobLigne);
        	  end;
         Ferme (QConso);
         end;
	   end;
  L.AllSelected:=False;
  //
  if GetParamSoc('SO_BTFACAPPELDETAIL') then
     CreerPiecesFromLignes(TobLigne,'FACDETAILCONSO', StrToDateTime(fDateFact.text), true, false, TobResult)
  Else
     CreerPiecesFromLignes(TobLigne,'FACCONSO',StrToDateTime(fDateFact.text), true, false, TobResult);
  //
  MajConsoLigne(TOBResult);
  MajEtatAffaire(TobResult);

	FINALLY
    SourisNormale;
    F.BChercheClick(ecran);
    TobLigne.free;
    TobResult.Free;
    TobConso.Free;
  END;

end;

Procedure TOF_BTFACCONSO.MajConsoLigne(TobResult : Tob);
Var j 		 : integer;
    i 		 : integer;
    Req		 : String;
    Num 	 : Integer;
    NumLig : Integer;
    NumMov : Integer;
Begin

    For i := 0 to TobResult.detail.count-1 do
    		Begin
        For j := 0 To Tobresult.Detail[I].detail.Count-1 Do
        		Begin
            if Tobresult.detail[i].detail[j].GetValue('GL_ARTICLE') <> '' then
               Begin
               Num := Tobresult.detail[i].detail[j].GetValue('GL_NUMERO');
               NumLig := Tobresult.detail[i].detail[j].GetValue('GL_NUMLIGNE');
               NumMov := Tobresult.detail[i].detail[j].GetValue('BLP_NUMMOUV');
               Req := 'UPDATE CONSOMMATIONS SET BCO_FACTURABLE="F';
               Req := Req + '", BCO_NATUREPIECEG="'+ Tobresult.detail[i].detail[j].GetValue('GL_NATUREPIECEG');
               Req := Req + '", BCO_SOUCHE="' + Tobresult.detail[i].detail[j].GetValue('GL_SOUCHE');
               Req := Req + '", BCO_NUMERO=' + IntToStr(Num);
               Req := Req + ', BCO_NUMORDRE=' + IntToStr(NumLig);
               Req := Req + ' WHERE BCO_NUMMOUV=' + IntToStr(NumMov);
               ExecuteSQL(Req);
               end;
		        end;
        end;

end;

Procedure TOF_BTFACCONSO.MajEtatAffaire(TobResult : TOB);
var NumAff : String;
    i 		 : integer;
    J 		 : integer;
Begin
	For i := 0 to TobResult.detail.count-1 do
  		Begin
      For j := 0 To Tobresult.Detail[I].detail.Count-1 Do
        	Begin
          NumAff := Tobresult.detail[i].detail[j].GetValue('GL_AFFAIRE');
      // s'il existe un devise non facturé sur cet appel on ne positionne pas comme facturé ou cloturé
      if ExisteSql ('SELECT 1 FROM PIECE WHERE GP_NATUREPIECEG="DAP" AND '+
      'GP_AFFAIRE="'+NumAff+'" AND '+
      'GP_DEVENIRPIECE = ""') then continue;
        //FV1 : 01/12/2015 - FS#1811 - TEAM RESEAUX : interventions - ajout option clôturé si facturé selon paramètre société.
        if GetParamSocSecur('SO_FACTCLOTURE', false) then
      begin
         if (Copy(NumAff,0,1) = 'W') then PositionneEtatAffaire(NumAff, 'CL1')
         else if (Copy(NumAff,0,1) = 'I') then PositionneEtatAffaire(NumAff, 'CLO')
         else if (Copy(NumAff,0,1) = 'A') then PositionneEtatAffaire(NumAff, 'CLO')
      end
        else
      begin
        if (Copy(NumAff,0,1) = 'W') then PositionneEtatAffaire(NumAff, 'FAC');
      end;
          End;
      End;

End;

Procedure TOF_BTFACCONSO.ChargementConsoLigne(TobConso, TOBPiece, TobLigne : Tob);
var TOBP	: TOB;
    TOBT	: TOB;
    TOBA	: TOB;
    TOBL	: TOB;
    TOBART: TOB;
    CleDoc:R_CLEDOC;
    QArt : TQuery;
    DEV : RDevise;
begin

  TOBT := TOB.create ('TIERS',nil,-1);
  TOBA := TOB.Create ('AFFAIRE',nil,-1);
	TOBART := CreerTOBArt (nil);

  //--Chargement de CleDoc
  FillChar(CleDoc,Sizeof(CleDoc),#0);

  if Getparamsoc('SO_FACTPROV') then
     CleDoc.NaturePiece:='FPR'
  else
     CleDoc.NaturePiece:='FAC';

  //FV - FS#421 - date facturation
  //CleDoc.DatePiece:=V_PGI.DateEntree;
  CleDoc.DatePiece:= StrToDateTime(fDateFact.text);

  CleDoc.Souche:=GetSoucheG(CleDoc.NaturePiece,'','') ;
  CleDoc.Indice:=0 ;

  //--Chargement du tiers
  TOBT.putValue('T_AUXILIAIRE',TOBConso.getValue('T_AUXILIAIRE'));
  TOBT.loadDb;

  //--Chargement de l'affaire
  if TOBConso.GetValue('BCO_AFFAIRESAISIE') = '' then
 		 TOBA.PutValue('AFF_AFFAIRE',TOBConso.GetValue('BCO_AFFAIRE'))
  Else
 	   TOBA.PutValue('AFF_AFFAIRE',TOBConso.GetValue('BCO_AFFAIRESAISIE'));

  TOBA.loadDb;

  //--Chargement de l'entete
  TOBP := TOBPiece.FindFirst(['GP_TIERS', 'GP_AFFAIRE'], [TOBT.GetValue('T_TIERS'),TOBA.GetValue('AFF_AFFAIRE')], True);
  if TOBP = nil then
   	 TOBP := CreerTOBPieceVide(Cledoc,TOBT.GetValue('T_TIERS'),TOBConso.GetValue('BCO_AFFAIRE'),'','',True,false);
  TOBL := NewTOBLigne (TOBLigne,0);
//  TOBL := TOB.Create ('LIGNE',TOBLigne,-1);
//  AddLesSupLigne (TOBL,false);
  PieceVersLigne (TOBP,TOBL);
  TOBL.clearDetail;
  //--Chargement des Articles
  QArt := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE FROM ARTICLE A '+
                 'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                 'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+
                 	TOBConso.getValue('BCO_ARTICLE')+'"',true);
  if not QArt.eof then
  begin
  	ChargerTobArt(TOBArt,nil,'','',QArt) ;
  end;
  ferme (QArt);

  DEV.Code := TOBP.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TOBP.getValue('GP_TAUXDEV');

  TOBL.PutValue('GL_REFARTSAISIE',TOBART.getValue('GA_CODEARTICLE'));

	ArticleVersLigne (TOBP,TOBArt,nil,TOBL,TOBT,false);
  ArticleDetailVersLigne(TOBL.getValue('GL_REFARTSAISIE'),TOBart,TOBL, [tModeGridStd]);
  TOBL.SetString('NUMBON',TOBA.GetString('AFF_NUMBON'));
  //--Chargement date de l'intervention dans la date de livraison
  // afin de la récupérer dnas la facture
  TOBL.PutValue('GL_DATELIVRAISON', TobConso.GetValue('BCO_DATEMOUV'));

  //--Chargement quantité et libellé de la ligne a facturer
  TOBL.PutValue('GL_QTEFACT', TobConso.GetValue('BCO_QTEFACTUREE'));
  TOBL.PutValue('GL_LIBELLE', TobConso.GetValue('BCO_LIBELLE'));
  TOBL.PutValue('GL_DPA', TobConso.GetValue('BCO_DPA'));
  TOBL.PutValue('GL_DPR', TobConso.GetValue('BCO_DPR'));
  TOBL.PutValue('GL_PUHT', TobConso.GetValue('BCO_PUHT'));
  TOBL.PutValue('GL_PUHTDEV', pivottodevise(TOBL.GetDouble('GL_PUHT'),DEV.Taux,DEV.Quotite,DEV.Decimale));
  TOBL.PutValue('GL_COEFFG', 0);
  TOBL.PutValue('GL_COEFFC', 0);
  TOBL.PutValue('GL_COEFFR', 0);
  if TobConso.GetValue('BCO_DPA') <> 0 then
  begin
  	TOBL.PutValue('GL_COEFFG', Arrondi((TobConso.GetValue('BCO_DPR')/TobConso.GetValue('BCO_DPA'))-1,4));
  end else
  begin
  	TOBL.PutValue('GL_COEFFG', 0);
  end;
  if TobConso.GetValue('BCO_DPR') <> 0 then
  begin
  	TOBL.PutValue('GL_COEFMARG', Arrondi(TobConso.GetValue('BCO_PUHT')/TobConso.GetValue('BCO_DPR'),4));
  end else
  begin
  	TOBL.PutValue('GL_COEFMARG', 1);
  end;
  CalculeLigneAc (TOBL,DEV,False);
  //
  IF TobConso.GetString('BCO_FAMILLETAXE1')<>'' then
  begin
    TOBL.SetString('GL_FAMILLETAXE1',TobConso.GetString('BCO_FAMILLETAXE1'));
  end;
  TOBL.PutValue('BLP_NUMMOUV', TobConso.GetValue('BCO_NUMMOUV'));
  TOBL.PutValue('BLP_INDICECON', TobConso.GetValue('BCO_INDICE'));

  //Chargement du code chantier saisit sur l'appel associé à la conso...
  Qart := Opensql('SELECT AFF_CHANTIER FROM AFFAIRE WHERE AFF_AFFAIRE ="' + TOBA.GetValue('AFF_AFFAIRE') + '"', true);
  If not Qart.eof then
     TOBL.PutValue('GL_CHANTIER', QArt.FindField('AFF_CHANTIER').AsString);
  ferme(QArt);

end;

Initialization
  registerclasses ( [ TOF_BTFACCONSO ] ) ;
end.

