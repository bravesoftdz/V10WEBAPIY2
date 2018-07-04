{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/11/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTAPPELCONSO ()
Mots clefs ... : TOF;BTAPPELCONSO
*****************************************************************}
Unit UTofAppelConso ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     HTB97,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     UtilSaisieConso,
     UTOB;

Type
  TOF_BTAPPELCONSO = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Argument : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private

		TobAppel		: Tob;
		TobConso		: Tob;

	  TobHeures   : Tob;
    TobFrais	  : Tob;
    TobPrestExt : Tob;
	  TobEngin    : Tob;
    TobMateriaux: Tob;
    TobRecettes : Tob;

    NoAppel			: String;

    GConso 			: THGrid;

    THEURES     : TTabSheet;
    TFRAIS      : TTabSheet;
    TMATERIELS	: TTabSheet;
    TFOURNITURES: TTabSheet;
    TMOEXT			: TTabSheet;
    TRECETTES 	: TTabSheet;

    PGTemps 		: TPageControl;
    BBlocNote 	: TToolbarButton97;
    PlusWindow 	: TToolWindow97;

    //Procedures liées aux évènement d'un objet
    procedure BBlocNoteClick(Sender: TOBject);
    procedure ChangeOnglet(Sender: TObject);

    //Declaration Procédure
    procedure ControleChamp(Champ, Valeur: String);
    procedure ControleCritere(Critere: String);
    procedure EventGrilles(Active: boolean);
    procedure PositionneDansGrid(TheGrid: THGrid; Arow, Acol: integer);
    procedure ShowToolWindow(Actif: boolean);
    
  end ;

Implementation

procedure TOF_BTAPPELCONSO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTAPPELCONSO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTAPPELCONSO.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTAPPELCONSO.OnLoad ;
Var QAppel	: TQuery;
		QQ			: TQuery;
		QConso	: TQuery;
    Req			: String;
begin
  Inherited ;

	//chargement de l'entête de l'Appel
  Req := '';
  Req := 'SELECT AFF_AFFAIRE, AFF_LIBELLE, AFF_DATEDEBUT, AFF_DATEFIN, AFF_RESPONSABLE ';
  Req := Req + ' FROM AFFAIRE WHERE AFF_AFFAIRE ="' + NoAppel + '"';
  QAppel := OpenSQL (Req, True);
  if QAppel.eof then
     Begin
     Ferme(QAppel);
     ecran.Close;
     end;

  ToBAppel := Tob.Create('AFFAIRE', nil, -1);
	TobAppel.selectDB('',QAppel);

  Ferme(QAppel);

  //Chargement des zones entete ecran avec les zones de la Tob
  SetControlText('AFF_AFFAIRE', NoAppel);
	SetControlText('AFF_LIBELLE', TobAppel.Getvalue('AFF_LIBELLE'));
  SetControltext('AFF_DATEDEBUT', DateToStr(TobAppel.GetValue('AFF_DATEDEBUT')));
  SetControltext('AFF_DATEFIN', DatetoStr(TobAppel.GetValue('AFF_DATEFIN')));
  SetControlText('AFF_HEUREDEBUT', TimeToStr(TobAppel.GetValue('AFF_DATEDEBUT')));
  SetControlText('AFF_HEUREFIN', TimeToStr(TobAppel.GetValue('AFF_DATEFIN')));
	SetControlText('AFF_RESSOURCE', TobAppel.GetValue('AFF_RESPONSABLE'));

  //Recherche du nom du responsable principal
  Req := '';
  Req := 'SELECT ARS_LIBELLE, ARS_LIBELLE2 FROM RESSOURCE ';
  Req := Req + 'WHERE ARS_RESSOURCE ="' + GetControlText('AFF_RESSOURCE') + '"';

  QQ := OpenSQL (Req, True);
  if QQ.eof then
		 SetControlText('AFF_RESSOURCE', '')
  Else
     SetControlText('AFF_RESSOURCE', QQ.findfield('ARS_LIBELLE').asString + ' ' + QQ.findfield('ARS_LIBELLE2').asString);

  Ferme(QQ);

  //chargement de la grille des appels
  Req:= '';
  Req := 'SELECT * FROM CONSOMMATIONS WHERE BCO_AFFAIRE ="' + NoAppel + '" ORDER BY BCO_NATUREMOUV';
	QConso := OpenSQL (Req, True);
  if QConso.eof then
  	 Begin
  	 Ferme(QConso);
     ecran.Close;
     end;

	//Chargement de la Tob des consommations
  TobConso := Tob.Create('CONSOMMATIONS', nil, -1);
  TobConso.LoadDetailDB('CONSOMMATIONS', '', 'BCO_NATUREMOUV',QConso, True);

  Ferme(QConso);

  OnDisplay;

end ;

procedure TOF_BTAPPELCONSO.OnArgument (Argument : String ) ;
Var Critere : String;
    Valeur  : String;
    Champ   : String;
    X       : integer;
begin
  Inherited ;

  //Récupération valeur de argument
	Critere:=(Trim(ReadTokenSt(Argument)));

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
      Critere := (Trim(ReadTokenSt(Argument)));
    end;

  GConso := THGrid(GetControl('GConso'));
  //
  THEURES := TTabSheet (GetControl('THEURES'));
  TFRAIS := TTabSheet (GetControl('TFRAIS'));
  TMATERIELS := TTabSheet (GetControl('TMATERIELS'));
  TFOURNITURES := TTabSheet (GetControl('TFOURNITURES'));
  TMOEXT := TTabSheet (GetControl('TMOEXT'));
  TRECETTES := TTabSheet (GetControl('TRECETTES'));
  //
  PGTemps := TPageControl (GetControl('PGTEMPS'));
  PGTemps.OnChange := ChangeOnglet;
  //
  BBLocNote := TToolbarButton97 (getControl('BBLOCNOTE'));
  BBlocNote.OnClick := BBlocNoteClick;
  //

end ;

procedure TOF_BTAPPELCONSO.ChangeOnglet (Sender : TObject);
var Cancel : boolean;
		Arow,Acol : integer;
begin
end;

Procedure TOF_BTAPPELCONSO.ControleChamp(Champ : String;Valeur : String);
Begin

	if Champ = 'APPEL' Then NoAppel := Valeur;

end;

procedure Tof_BTAPPELCONSO.BBlocNoteClick (Sender : TOBject);
begin
  if not (Sender Is TToolbarButton97) then exit;

  ShowToolWindow(TToolBarButton97(Sender).Down);

end;

procedure Tof_BTAPPELCONSO.ShowToolWindow(Actif : boolean);
begin
  if Actif Then
     PlusWindow.Visible := true
  else
     PlusWindow.Visible := false;
end;


Procedure Tof_BTAPPELCONSO.ControleCritere(Critere : String);
Begin
end;


procedure TOF_BTAPPELCONSO.OnClose ;
begin
  Inherited ;

  TobAppel.Free;
  TobConso.Free;

  TobHeures.Free;
  TobFrais.Free;
  TobPrestExt.Free;
	TobEngin.Free;
  TobMateriaux.Free;
  TobRecettes.Free;
                  
end ;

procedure TOF_BTAPPELCONSO.OnDisplay () ;
var Arow   	: Integer;
    Acol 		: Integer;
    Cancel 	: Boolean;
    I				: Integer;
    J				: Integer;
    ZoneGrille	: String;
begin
  Inherited ;

  TobHeures := Tob.create('Heures',nil, -1);
  TobFrais := Tob.create('Frais',nil, -1);
  TobPrestExt := Tob.create('PrestExt',nil, -1);
  TobEngin := Tob.create('Engin',nil, -1);
  TobMateriaux := Tob.create('Materiaux',nil, -1);
  TobRecettes := Tob.create('Recettes',nil, -1);

	//Remise à blanc des tob d'affichage
  TobHeures.ClearDetail;
  TobFrais.ClearDetail;
  TobPrestExt.ClearDetail;
	TobEngin.ClearDetail;
  TobMateriaux.ClearDetail;
  TobRecettes.ClearDetail;

  ZoneGrille := 'BCO_DATEMOUV;BCO_RESSOURCE;BCO_ARTICLE;BCO_TYPEHEURE;';
  ZoneGrille := ZoneGrille + 'BCO_LIBELLE; BCO_QUANTITE; BCO_QUALIFQTEMOUV; BCO_DPR; BCO_MONTANTPR';

  //Chargement des Tobs d'affichage en fonction des Onglets
  For I := 0 to TobConso.Detail.count-1 Do
  		Begin
      If TobConso.Detail[I].GetValue('BCO_NATUREMOUV') = 'MO' Then
         Begin
         TobConso.PutLigneGrid(GConso,I,False, False, ZoneGrille);
         //AddLesSupLignesConso(TobHeures);
         end
      Else if TobConso.Detail[i].GetValue('BCO_NATUREMOUV') = 'FRS' Then
         Begin
         //TobConso.PutGridDetail(GConso,True, True, ZoneGrille);
         //AddLesSupLignesConso(TobFrais);
         End
      Else if TobConso.Detail[i].GetValue('BCO_NATUREMOUV') = 'EXT' Then
         Begin
         //AddLesSupLignesConso(TobPrestExt);
         End
      Else if TobConso.Detail[i].GetValue('BCO_NATUREMOUV') = 'RES' Then
         Begin
         //AddLesSupLignesConso(TobEngin);
         End
      Else if TobConso.Detail[i].GetValue('BCO_NATUREMOUV') = 'FOU' Then
         Begin
         //AddLesSupLignesConso(TobMateriaux);
         end
      Else if TobConso.Detail[i].GetValue('BCO_NATUREMOUV') = 'RAN' Then
         Begin
         //AddLesSupLignesConso(TobRecettes);
         End
      Else if TobConso.Detail[i].GetValue('BCO_NATUREMOUV') = 'FAN' Then
         Begin
         //AddLesSupLignesConso(TobRecettes);
         End;
      End;

  //DefiniGrilles;
  //Remplitgrilles;
  //EventGrilles(true);

    Arow := 1;
    Acol := 1;
    PositionneDansGrid (GConso,Arow,Acol);

    //ActiveMode := TgsMO;
    PGTemps.ActivePage := THeures;

end ;


procedure TOF_BTAPPELCONSO.PositionneDansGrid (TheGrid : THGrid;Arow,Acol : integer);
var cancel : boolean;
begin

  TheGrid.row := Arow;
  TheGrid.col := Acol;
	TheGrid.ShowEditor;

end;

procedure TOF_BTAPPELCONSO.EventGrilles(Active : boolean);
begin
  if Active then
	  begin
    // grilles de saisie des heure de MO
    //GConso.OnRowEnter := GConsoRowEnter;
    //GConso.OnRowExit := GConsoRowExit;
    //GConso.OnCellEnter  := GConsoCellEnter;
    //GConso.OnCellExit  := GConsoCellExit;
    //GConso.OnElipsisClick := GConsoElipsisClick;
    //GConso.OnKeyDown := GridKeyDown;
    //GConso.PostDrawCell := GConsoPostDrawCell;
  	end
  else
  	begin
    // grilles de saisie des heure de MO
    GConso.OnRowEnter := nil;
    GConso.OnRowExit := nil;
    GConso.OnCellEnter  := nil;
    GConso.OnCellExit  := nil;
    GConso.OnElipsisClick := nil;
    GConso.PostDrawCell := nil;
  	end;
// --
end;

procedure TOF_BTAPPELCONSO.OnCancel () ;
begin
  Inherited ;
end ;


Initialization
  registerclasses ( [ TOF_BTAPPELCONSO ] ) ;
end.

