{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 07/08/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RECALCPIECE_RAP ()
Mots clefs ... : TOF;RECALCPIECE_RAP
*****************************************************************}
Unit RECALCPIECE_RAP_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
		 MaineAGL,
     eMul,
{$ENDIF}
     forms,
     UTOB,
     AglInit,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Splash,
     PiecesRecalculs,
     FermeturePieces,
     EpurationPieces,
     EpurationChantiers,
     UTOF ;

Type
  TOF_RECALCPIECE_RAP = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	TRAITEMANPIECE, TRAITFERMETURE,TRAITEPURATION,TRAITEEPUCHANTIER,AFFICHEERREURS : boolean;
  	TOBPieces : TOB;
    Memo : Tmemo;
    CalculPv : boolean;
    fermeture : boolean;
    procedure AfficheInfoTraitementPiece (TOBPiece : TOB; RetourTrait : TRResult );
    procedure AfficheInfoFermeturePiece (TOBPiece : TOB; retourFermeture : integer);
    procedure AfficheInfoEpurationPiece (TOBpiece : TOB ;RetourEpuration : integer);
		procedure AfficheInfoEpurationChantier (TOBChantier : TOB; RetourEpuration : integer);
    //
    procedure EnregistreEventFermeture;
    procedure EnregistreEventEpuration;
		procedure EnregistreEventEpurationChantier;
    procedure EnregistreEvent(TypeEvt, StError: string);
  end ;

procedure TraiteRecalculPieces (TOBPieces : TOB; CalculPv,ReajustePVOuv : boolean);
Procedure TraiteManPieces (TOBPiece : TOB; Silent : boolean=false);
procedure TraiteFermeturePieces (TOBPieces : TOB;Ferme:boolean);
procedure TraiteEpurePieces (TOBPieces: TOB);
procedure TraiteEpureChantiers (TOBChantiers : TOB);
procedure RapportsErreurs (TOBErreurs : TOB);

Implementation
var Splash : TFsplashScreen;

procedure RapportsErreurs (TOBErreurs : TOB);
begin
	TheTOB := TOBErreurs;
  AglLanceFiche ('BTP','BTRECALCPIECE_RAP','','','ERREURS');
  TheTOB := Nil;
end;

procedure TraiteEpureChantiers (TOBChantiers : TOB);
begin
  //
  Splash := TFsplashScreen.Create (application);
  Splash.Label1.Caption  := TraduireMemoire('Epuration des chantiers en cours...');
  Splash.Show;
  Splash.BringToFront;
  Splash.refresh;
	TheTOB := TOBChantiers;
  AglLanceFiche ('BTP','BTRECALCPIECE_RAP','','','EPUCHANTIER');
  TheTOB := Nil;
  Splash.free;
  Splash := nil;
end;

procedure TraiteEpurePieces (TOBPieces: TOB);
var Splash : TFsplashScreen;
begin
  //
  Splash := TFsplashScreen.Create (application);
  Splash.Label1.Caption  := TraduireMemoire('Epuration des pièces en cours...');
  Splash.Show;
  Splash.BringToFront;
  Splash.refresh;
	TheTOB := TOBPieces;
  AglLanceFiche ('BTP','BTRECALCPIECE_RAP','','','EPURATION');
  TheTOB := Nil;
  Splash.free;
end;

procedure TraiteFermeturePieces (TOBPieces : TOB;Ferme:boolean);
var Splash : TFsplashScreen;
		Libmess : string;
begin
  //
  if Ferme then Libmess := 'réouverture' else Libmess := 'fermeture';
  Splash := TFsplashScreen.Create (application);
  Splash.Label1.Caption  := TraduireMemoire(Libmess + ' des pièces en cours...');
  Splash.Show;
  Splash.BringToFront;
  Splash.refresh;
	TheTOB := TOBPieces;
  if Ferme then TheTOB.AddChampSupValeur ('FERMETURE','-')
  				 else TheTOB.AddChampSupValeur ('FERMETURE','X');
  AglLanceFiche ('BTP','BTRECALCPIECE_RAP','','','FERMETURE');
  TheTOB := Nil;
  Splash.free;
end;

procedure TraiteRecalculPieces (TOBPieces : TOB; CalculPv,ReajustePVOuv : boolean);
var Splash : TFsplashScreen;
begin
  //
  Splash := TFsplashScreen.Create (application);
  Splash.Label1.Caption  := TraduireMemoire('Recalcul des pièces en cours...');
  Splash.Show;
  Splash.BringToFront;
  Splash.refresh;
	TheTOB := TOBPieces;
  if CalculPv then TheTOB.AddChampSupValeur ('CALCULPV','X')
  						else TheTOB.AddChampSupValeur ('CALCULPV','-');
	if ReajustePVOuv then TheTOB.AddChampSupValeur ('RAJUSTEPVOUV','X')
  								 else TheTOB.AddChampSupValeur ('RAJUSTEPVOUV','-');
  AglLanceFiche ('BTP','BTRECALCPIECE_RAP','','','');
  TheTOB := Nil;
  Splash.free;
end;

Procedure TraiteManPieces (TOBPiece : TOB; Silent : boolean=false);
var Splash  : TFsplashScreen;
    iOuv    : Integer;
    iPiece  : Integer;
    iIndOuv : Integer;
    NumPiece: Integer;
    TheText : string;
    Memo    : TMemo;
begin
  //
  Splash := TFsplashScreen.Create (application);
  Splash.Label1.Caption  := TraduireMemoire('Mise à niveau des pièces en cours...');
  Splash.Show;
  Splash.BringToFront;
  Splash.refresh;

  TheTOB := TOBPiece;
  //
  for iPiece := 0 to Tobpiece.detail.Count - 1  do
  begin
    iIndOuv := 0;
    NumPiece := TOBPiece.detail[iPiece].GetInteger('GP_NUMERO');
    for iOuv := 0 to TOBPiece.Detail[iPiece].Detail.Count - 1 do
    begin
      Inc(iIndOuv);
      TOBPiece.detail[iPiece].Detail[iOuv].PutValue('BLO_UNIQUEBLO', iIndOuv);
      TOBPiece.detail[iPiece].Detail[iOuv].updateDb;
    end;
    TobPiece.detail[iPiece].PutValue('GP_UNIQUEBLO', iIndOuv);
    TOBPiece.detail[iPiece].updateDb;
    TheText := Format ('%s numéro %d Mise à niveau',
                       [Rechdom('GCNATUREPIECEG',TOBPiece.Detail[iPiece].GetString('GP_NATUREPIECEG'),false),NumPiece]);
    TOBPiece.detail[ipiece].AddChampSupValeur('MEMO', TheText);
  end;
  
	if not Silent then AglLanceFiche ('BTP','BTRECALCPIECE_RAP','','','TRAITEMANPIECE');

  TheTOB := Nil;

  Splash.free;

end;

procedure TOF_RECALCPIECE_RAP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RECALCPIECE_RAP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RECALCPIECE_RAP.EnregistreEvent(TypeEvt, StError : string);
var QQ      : TQuery;
    NumEvent: integer;
    TOBE    : TOB;
begin

  NumEvent := 0;
  QQ := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True,-1, '', True);
  if not QQ.EOF then NumEvent := QQ.Fields[0].AsInteger;
  Ferme(QQ);
  Inc(NumEvent);
  TOBE := TOB.Create ('JNALEVENT',nil,-1);
  TOBE.PutValue('GEV_NUMEVENT', NumEvent);
  TOBE.PutValue('GEV_TYPEEVENT',TypeEvt);
  TOBE.PutValue('GEV_LIBELLE',Copy(ecran.caption, 1, 35));
  TOBE.PutValue('GEV_DATEEVENT',Date);
  TOBE.PutValue('GEV_UTILISATEUR',V_PGI.User);
  TOBE.PutValue('GEV_ETATEVENT','OK');
  TOBE.PutValue('GEV_BLOCNOTE',TMemo(GetControl('RAPPORT')).Lines.Text);
  if not TOBE.InsertDB(nil) then PgiInfo (StError);
end;


procedure TOF_RECALCPIECE_RAP.EnregistreEventFermeture;
var QQ: TQuery;
  NumEvent: integer;
  TOBE : TOB;
begin

  NumEvent := 0;
  QQ := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True,-1, '', True);
  if not QQ.EOF then NumEvent := QQ.Fields[0].AsInteger;
  Ferme(QQ);
  Inc(NumEvent);
  TOBE := TOB.Create ('JNALEVENT',nil,-1);
  TOBE.PutValue('GEV_NUMEVENT', NumEvent);
  TOBE.PutValue('GEV_TYPEEVENT','BFP');
  TOBE.PutValue('GEV_LIBELLE',Copy(ecran.caption, 1, 35));
  TOBE.PutValue('GEV_DATEEVENT',Date);
  TOBE.PutValue('GEV_UTILISATEUR',V_PGI.User);
  TOBE.PutValue('GEV_ETATEVENT','OK');
  TOBE.PutValue('GEV_BLOCNOTE',TMemo(GetControl('RAPPORT')).Lines.Text);
  if not TOBE.InsertDB(nil) then PgiInfo ('Erreur ecriture Mvts');
end;

procedure TOF_RECALCPIECE_RAP.EnregistreEventEpuration;
var QQ: TQuery;
  NumEvent: integer;
  TOBE : TOB;
begin

  NumEvent := 0;
  QQ := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True,-1, '', True);
  if not QQ.EOF then NumEvent := QQ.Fields[0].AsInteger;
  Ferme(QQ);
  Inc(NumEvent);
  TOBE := TOB.Create ('JNALEVENT',nil,-1);
  TOBE.PutValue('GEV_NUMEVENT', NumEvent);
  TOBE.PutValue('GEV_TYPEEVENT','EPU');
  TOBE.PutValue('GEV_LIBELLE',Copy(ecran.caption, 1, 35));
  TOBE.PutValue('GEV_DATEEVENT',Date);
  TOBE.PutValue('GEV_UTILISATEUR',V_PGI.User);
  TOBE.PutValue('GEV_ETATEVENT','OK');
  TOBE.PutValue('GEV_BLOCNOTE',TMemo(GetCOntrol('RAPPORT')).Lines.Text);
  if not TOBE.InsertDB(nil) then PgiInfo ('Erreur ecriture Mvts');
end;

procedure TOF_RECALCPIECE_RAP.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_RECALCPIECE_RAP.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_RECALCPIECE_RAP.OnArgument (S : String ) ;
var Indice : integer;
		retourTrait : TRResult;
    RetourFermeture,RetourEpuration,RetourEpurationChantier : integer;
    RetourMiseANiveau : Integer;
    ReajustePvPouv,AvecControle : Boolean;
    Debut:TDateTime;
begin
  Inherited ;
  AFFICHEERREURS := false;
  TOBPieces := LaTOB;
  if S = 'ERREURSOUPAS' then
  begin
    AFFICHEERREURS := true;
    Ecran.Caption := 'Rapport de Génération';
  end else if S = 'ERREURS' then
  begin
    AFFICHEERREURS := true;
    Ecran.Caption := 'Rapport d''anomalies';
  end else
  begin
    if (Pos('FERMETURE',S)=0) and (pos('EPURATION',S)=0) and (pos('EPUCHANTIER',S)=0) and (Pos('TRAITEMANPIECE', S)=0) then
    begin
      TRAITEMANPIECE := False;
      TRAITFERMETURE := false;
      TRAITEPURATION := false;
      TRAITEEPUCHANTIER := false;
      CalculPv := (TOBPieces.getValue('CALCULPV')='X');
      ReajustePvPouv := (TOBPieces.getValue('RAJUSTEPVOUV')='X');
    end else if (pos('EPURATION',S)>0) then
    begin
      TRAITEMANPIECE := False;
      TRAITFERMETURE := false;
      TRAITEEPUCHANTIER := false;
      TRAITEPURATION := true;
      ecran.Caption := 'Epuration de pièces';
    end else if (pos('EPUCHANTIER',S)>0) then
    begin
      AvecControle := StrToBool_(TOBPieces.GetString('AVECCONTROLE'));
      TRAITEMANPIECE := False;
      TRAITFERMETURE := false;
      TRAITEEPUCHANTIER := true;
      TRAITEPURATION := false;
      ecran.Caption := 'Epuration de chantiers';
    end else if (pos('TRAITEMANPIECE',S)>0) then
    begin
      TRAITEMANPIECE := True;
      TRAITFERMETURE := false;
      TRAITEEPUCHANTIER := false;
      TRAITEPURATION := false;
      ecran.Caption := 'Mise à niveau des Documents';
    end else
    begin
      TRAITEMANPIECE := False;
      TRAITFERMETURE := True;
      TRAITEPURATION := false;
      TRAITEEPUCHANTIER := false;
      fermeture := (TOBPieces.getValue('FERMETURE')='X');
      if fermeture then ecran.Caption := 'Fermeture de pièces' else ecran.Caption := 'Réouverture de pièces';
    end;
  end;
  Memo := Tmemo(GetControl('RAPPORT'));

  if AFFICHEERREURS then
  begin
		for Indice := 0 to TOBPieces.detail.count -1 do
    begin
    	Memo.Lines.Add (TOBPieces.detail[Indice].GetString('RAPPORT'));
    end;
  end else begin
    if TRAITEEPUCHANTIER then
    begin
    	Memo.Lines.Add ('-------------------------------------------');
    	Memo.Lines.Add ('Début le '+DateTimeToStr(Now));
    	Memo.Lines.Add ('-------------------------------------------');
    end;
    For Indice := 0 to TOBPieces.detail.count -1 do
    begin
      if (not TRAITFERMETURE) and (not TRAITEPURATION) and (not TRAITEEPUCHANTIER) and (not TRAITEMANPIECE) then
      begin
        retourTrait := TraitementRecalculPiece (TOBPieces.detail[Indice],CalculPv,ReajustePvPouv);
        AfficheInfoTraitementPiece (TOBPieces.detail[Indice],RetourTrait);
      end else if TRAITEPURATION then
      begin
        RetourEpuration := TraitementEpurationPiece (TOBPieces.detail[Indice]);
        AfficheInfoEpurationPiece (TOBPieces.detail[Indice],RetourEpuration);
      end else if TRAITEEPUCHANTIER then
      begin
        RetourEpurationChantier := TraitementEpurationChantier (TOBPieces.detail[Indice],AvecControle,Splash);
        AfficheInfoEpurationChantier (TOBPieces.detail[Indice],RetourEpurationChantier);
      end else if TRAITEMANPIECE then
      begin
        Memo.Lines.Add(TOBPieces.detail[indice].GetString('MEMO'));
      end else
      begin
        RetourFermeture := TraitementFermeturePiece (TOBPieces.detail[Indice],Fermeture);
        AfficheInfoFermeturePiece (TOBPieces.detail[Indice],RetourFermeture);
      end;
    end;
    if TRAITEEPUCHANTIER then
    begin
    	Memo.Lines.Add ('-------------------------------------------');
    	Memo.Lines.Add ('Fin le '+DateTimeToStr(Now));
    	Memo.Lines.Add ('-------------------------------------------');
    end;
    //
    if TRAITEMANPIECE then
    begin
      EnregistreEvent('BOU','Erreur ecriture Mvts')
    end else if TRAITFERMETURE Then
    begin
      EnregistreEventFermeture;
    end else if TRAITEPURATION then
    begin
      EnregistreEventEpuration;
    end else if TRAITEEPUCHANTIER then
    begin
      EnregistreEventEpurationChantier;
    end;
  end;

  //
end ;

procedure TOF_RECALCPIECE_RAP.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_RECALCPIECE_RAP.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_RECALCPIECE_RAP.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_RECALCPIECE_RAP.AfficheInfoTraitementPiece(TOBPiece: TOB; RetourTrait: TRResult);
var TheText : string;
		TheNumPiece : integer;
begin

	if ( not TRAITFERMETURE) then
  begin
    TheNumPiece := TOBPiece.getValue('GP_NUMERO');
    if RetourTrait = TrrOk then
    begin
    TheText := Format ('%s numéro %d calculé(e)',
                       [Rechdom('GCNATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'),false),TheNumPiece]);
    end else if RetourTrait = TrrErrBeforeCalc then
    begin
      TheText := Format ('Erreur sur  %s numéro %d Avant application des coeficients de Frais',
                         [Rechdom('GCNATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'),false),TheNumPiece]);
    end else if RetourTrait = TrrErrCalcPR then
    begin
      TheText := Format ('Erreur sur  %s numéro %d en application des coeficients de Frais',
                         [Rechdom('GCNATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'),false),TheNumPiece]);
    end else if RetourTrait = TrrErrCalcPV then
    begin
      TheText := Format ('Erreur sur  %s numéro %d en Calcul des PV',
                         [Rechdom('GCNATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'),false),TheNumPiece]);
    end else if RetourTrait = TrrErrEcriture then
    begin
      TheText := Format ('Erreur sur  %s numéro %d en validation',
                         [Rechdom('GCNATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'),false),TheNumPiece]);
    end;
  end;

  Memo.Lines.Add (TheText);

end;

procedure TOF_RECALCPIECE_RAP.AfficheInfoFermeturePiece(TOBPiece: TOB; retourFermeture: integer);
var TheText,LibTrt : string;
		TheNumPiece : integer;
begin
  TheNumPiece := TOBPiece.getValue('GP_NUMERO');
  if retourFermeture = 0 then
  begin
    if fermeture then Libtrt := 'fermé(e)' else Libtrt := 'réouvert(e)';
    TheText := Format ('%s numéro %d %s',
                       [Rechdom('GCNATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'),false),TheNumPiece,Libtrt]);
  end else if retourFermeture < 0 then
  begin
    TheText := Format ('Erreur écriture sur %s numéro %d ',
                       [Rechdom('GCNATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'),false),TheNumPiece]);
  end else if (retourFermeture = 1) then
  begin
    TheText := Format ('Impossible de réouvrir %s numéro %d. Le dossier de facturation à été cloturé',
                       [Rechdom('GCNATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'),false),TheNumPiece]);
  end else if retourFermeture = 2 then
  begin
    TheText := Format ('Impossible de réouvrir %s numéro %d. Le chantier à été cloturé',
                       [Rechdom('GCNATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'),false),TheNumPiece]);
  end else if retourFermeture = 11 then
  begin
    TheText := Format ('Impossible de réouvrir le(la) %s numéro %d.Vous ne pouvez réouvrir que la dernière situation',
                       [Rechdom('GCNATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'),false),TheNumPiece]);
  end else if retourFermeture = 20 then
  begin
    TheText := Format ('Impossible de réouvrir le(la) %s numéro %d.Le chantier est terminé',
                       [Rechdom('GCNATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'),false),TheNumPiece]);
  end;
  Memo.Lines.Add (TheText);
end;


procedure TOF_RECALCPIECE_RAP.AfficheInfoEpurationPiece(TOBpiece: TOB; RetourEpuration: integer);
var TheText,LibTrt : string;
		TheNumPiece : integer;
begin
  TheNumPiece := TOBPiece.getValue('GP_NUMERO');
  if RetourEpuration = 0 then
  begin
    Libtrt := 'épurée';
    TheText := Format ('%s numéro %d %s',
                       [Rechdom('GCNATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'),false),TheNumPiece,Libtrt]);
  end else
  begin
    Libtrt := 'non épurée';
    TheText := Format ('%s numéro %d %s',
                       [Rechdom('GCNATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'),false),TheNumPiece,Libtrt]);
  end;
  Memo.Lines.Add (TheText);
end;

procedure TOF_RECALCPIECE_RAP.EnregistreEventEpurationChantier;
var QQ: TQuery;
  NumEvent: integer;
  TOBE : TOB;
begin

  NumEvent := 0;
  QQ := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True,-1, '', True);
  if not QQ.EOF then NumEvent := QQ.Fields[0].AsInteger;
  Ferme(QQ);
  Inc(NumEvent);
  TOBE := TOB.Create ('JNALEVENT',nil,-1);
  TOBE.PutValue('GEV_NUMEVENT', NumEvent);
  TOBE.PutValue('GEV_TYPEEVENT','EPU');
  TOBE.PutValue('GEV_LIBELLE',Copy(ecran.caption, 1, 35));
  TOBE.PutValue('GEV_DATEEVENT',Date);
  TOBE.PutValue('GEV_UTILISATEUR',V_PGI.User);
  TOBE.PutValue('GEV_ETATEVENT','OK');
  TOBE.PutValue('GEV_BLOCNOTE',TMemo(GetCOntrol('RAPPORT')).Lines.Text);
  if not TOBE.InsertDB(nil) then PgiInfo ('Erreur ecriture Mvts');
end;

procedure TOF_RECALCPIECE_RAP.AfficheInfoEpurationChantier(
  TOBChantier: TOB; RetourEpuration: integer);
var TheChantier,TheText : string;
begin
  TheChantier := TOBChantier.getValue('AFF_AFFAIRE');
  if RetourEpuration = 0 then
  begin
    TheText := Format ('Chantier %s épuré',[TheChantier]);
  end else if RetourEpuration = 1 then
  begin
    TheText := Format ('Chantier %s non épuré: La cloture de facturation n''a pas été effectuée', [TheChantier]);
  end else
  begin
    TheText := Format ('Chantier %s non épuré: Erreur durant le traitement', [TheChantier]);
  end;
  Memo.Lines.Add (TheText);
end;

Initialization
  registerclasses ( [ TOF_RECALCPIECE_RAP ] ) ;
end.

