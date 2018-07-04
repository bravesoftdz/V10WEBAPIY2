{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 17/09/2002
Modifié le ... : 25/09/2002
Description .. : Source TOF de la FICHE : CPTVAEDI ()
Suite ........ : Procédure de Télé-TVA en EDI
Mots clefs ... : TOF;CPTVAEDI:TVA;EDI
*****************************************************************}
Unit UTOF_CPTVAEDI ;

Interface

Uses StdCtrls, Controls, Classes, Windows, Shellapi,
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox,
{$IFDEF EAGLCLIENT}
 	 MainEagl,
{$ELSE}
   db, dbtables, FE_main,
{$ENDIF}
   uTob, UTOF, paramSoc, AnnOutils, PGIEnv, LIA_COMMUN ;

procedure CPLanceFiche_TVAEDI;

Type
  TOF_CPTVAEDI = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
  	GrilleMontants : THGrid;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    function CoherenceMontants : Boolean;
    function CreerFichierEdiTVA : Boolean;
    procedure EcritureRib (var F :TextFile);
    procedure AccesControles;
    function GetCodePeriodicite : String;
  end ;

Implementation

Const MessageListe : Array[1..16] of String =		// Message d'erreurs
	{01}				('Vous devez renseigner un code FRP.',
  {02}         'Le code FRP est incorrect.',
  {03}         'Vous devez renseigner un numéro de SIRET.',
  {04}         'Le numéro de SIRET est incorrect.',
  {05}         'La somme des montants saisis ne correspond pas au montant total de TVA due. Voulez-vous continuer ?',
	{06}				 'Le numéro de TVA intracommunautaire est incorrect, veuillez le modifier.',
	{07}				 'Vous devez renseigner un numéro de SIRET du déclarant.',
  {08}         'Le numéro de SIRET du déclarant est incorrect.',
  {09}         'Les paramètres de CA3 ne sont pas renseignés.',
  {10}         'Il n''y a aucune déclaration de CA3 ouverte.',
  {11}         'Les comptes de banque associés à la procédure TELETVA dans les	 paramètres sociétés ne sont pas renseignés.',
	{12}				 'Les rubriques de la déclaration de TVA ne sont pas correctement renseignées.',
  {13}         'Le montant total de la déclaration de TVA en cours est nulle.',
  {14}         'Il n''y a aucune déclaration CA3 à transmettre.',
	{15}				 'Attention ! Vous devez installer TDI-Publifi',
  {16}				 'Certains montants sont mal renseignés, veuillez les modifier.'
 		          );

//==============================================================================
procedure CPLanceFiche_TVAEDI;
begin
	AGLLanceFiche('CP','CPTVAEDI','','','');
end;
//==============================================================================

procedure TOF_CPTVAEDI.OnUpdate ;
begin
  Inherited ;
	// Test du FRP
  if GetControlText('FRP')='' then
    begin
    PGIBox(TraduireMemoire(MessageListe[1]),Ecran.Caption);
    SetControlEnabled('FRP',True);
    SetFocusControl('FRP');
    LastError := 1;
    Exit;
    end
	else
    if not CoherenceFRP(GetControlText('FRP')) then
      begin
      PGIBox(TraduireMemoire(MessageListe[2]),Ecran.Caption);
	    SetControlEnabled('FRP',True);
  	  SetFocusControl('FRP');
      LastError := 1;
      Exit;
      end;

	// Test du SIRET
  if GetControlText('SIRET') = '' then
  	begin
    PGIBox(TraduireMemoire(MessageListe[3]),Ecran.Caption);
    SetControlEnabled('SIRET',True);
 	  SetFocusControl('SIRET');
    LastError := 1;
    Exit;
    end
  else
    if not CoherenceSIRET(GetControlText('SIRET'), length(GetControlText('SIRET'))) then
      begin
      PGIBox(TraduireMemoire(MessageListe[4]),Ecran.Caption);
	    SetControlEnabled('SIRET',True);
 		  SetFocusControl('SIRET');
      LastError := 1;
      Exit;
      end;

  if GetControlText('SIRETDECLA') = '' then
  	begin
    PGIBox(TraduireMemoire(MessageListe[7]),Ecran.Caption);
    LastError := 1;
    Exit;
    end
  else
    if not CoherenceSIRET(GetControlText('SIRETDECLA'), length(GetControlText('SIRETDECLA'))) then
      begin
      PGIBox(TraduireMemoire(MessageListe[8]),Ecran.Caption);
      LastError := 1;
      Exit;
      end;

	// Test sur les montants
  if not (IsNumeric(GrilleMontants.cells[4,1])
  				and IsNumeric(GrilleMontants.cells[4,1])
  				and IsNumeric(GrilleMontants.cells[4,1])) then
    begin
    PGIBox(TraduireMemoire(MessageListe[16]),Ecran.Caption);
    GrilleMontants.row := 1;
    GrilleMontants.col := 4;
    SetFocusControl('GRILLEMONTANTS');
    LastError := 1;
    Exit;
    end
	else
    if not CoherenceMontants then
      if PGIAsk(TraduireMemoire(MessageListe[5]),Ecran.Caption) <> mrYes then
        begin
        LastError := 1;
        Exit;
        end;

  // Génération du fichier
	CreerFichierEdiTVA;
end ;

procedure TOF_CPTVAEDI.OnLoad ;
var Q                         : TQuery;
    i                         : Integer;
    MontantTotal, cheminSource: String;
    F                         : TextFile;
    SLect                     : string;
begin
  Inherited ;
  // Init des champs avec ParamSoc
  SetControlText('RAISONSOCIALE',GetParamSocSecur('SO_LIBELLE',''));
  SetControlText('ADRESSE',GetParamSocSecur('SO_ADRESSE1',''));
  SetControlText('CODEPOSTAL',GetParamSocSecur('SO_CODEPOSTAL',''));
  SetControlText('VILLE',GetParamSocSecur('SO_VILLE',''));
  SetControlText('SIRET',GetParamSocSecur('SO_SIRET',''));
  SetcontrolText('PERIODE',DateToStr(iDate1900));
  SetcontrolText('PERIODE_',DateToStr(iDate1900));

  // Init des champs avec CA3_PARAMETRES
	Q := OpenSQL('SELECT * FROM CA3_PARAMETRE',True);
  if not Q.Eof then
  	begin
    // Init des champs par défaut
    SetcontrolText('PERIODE',Q.FindField('CP3_DATEENCOURS').asString);
    SetControlText('FRP',Q.FindField('CP3_RECETTE').asString
    										+Q.FindField('CP3_DOSSIER').asString
                        +Q.FindField('CP3_CLEDECLA').asString);
    SetControlText('CODEACTIVITE',Q.FindField('CP3_CODEACTIVE').asString);
    SetControlText('TVA',Q.FindField('CP3_INTRACOM').asString);
    end
  else
  	begin
    PGIINFO(TraduireMemoire(MessageListe[9]),Ecran.Caption);
    end;
  Ferme(Q);

  // Init des champs avec CA3_DECLARATION
  Q := OpenSql ('SELECT CA3_DATEFIN from CA3_DECLARATION Where CA3_STATUT="-"', TRUE);
  if not Q.EOF
  	then SetcontrolText('PERIODE_',Q.FindField ('CA3_DATEFIN').asString)
    else PGIINFO(TraduireMemoire(MessageListe[10]),Ecran.Caption);
  Ferme (Q);

  if GetParamSocSecur ('SO_TELETVABQE1','') = '' then
    PGIINFO(TraduireMemoire(MessageListe[11]) ,Ecran.Caption);

  // Infos déclarant dans fichier si existant
  if ctxPCL in V_PGI.PGIContexte
  	then CheminSource := V_PGI_ENV.pathDos + '\Tva.txt'
	  else CheminSource := V_PGI_ENV.RepLocal + '\STD\Tva.txt';
  if FileExists (CheminSource) then
  begin
    try
       AssignFile(F, CheminSource);
       Reset(F);
       while not EOF(F) do
       begin
            Readln(F, SLect);
            if Copy (SLect,1,5) = '00015'  then
              SetControlText('SIRETDECLA', Copy (SLect,7,length (SLect)- 6));

            if Copy (SLect,1,5) = '00016' then
            begin
              SetControlText('ENTITEDECLA', Copy (SLect,7,length (SLect)- 6));
              break;
            end;
       end;
    Finally
       CloseFile(F);
    end;
  end;

  // Montant total de TVA due issu de la déclaration :
  Q := OpenSql ('SELECT F10_VALRUB FROM F3310_VALRUB'
  							+ ' WHERE F10_DATEEFFET="' + USDateTime(LiasseEnv.DateEnCours) +'"'
								+ ' AND F10_MILLESIME = "' + LiasseEnv.MILLESIME + '"'
  							+ ' AND F10_NUMRUB="05035"' , TRUE);
  if not Q.EOF then
  	begin
		MontantTotal := Q.FindField('F10_VALRUB').asString;
		if Valeur(MontantTotal) = 0
  	  then PGIINFO(TraduireMemoire(MessageListe[13]),Ecran.Caption)
	    else THNumedit(GetControl('MONTANTTOTAL')).value := Valeur(MontantTotal);
    end
  else PGIINFO(TraduireMemoire(MessageListe[12]),Ecran.Caption);
  Ferme (Q);

  // Init des 3 comptes de banques
	For i:=1 to 3 do
    if GetParamSocSecur('SO_TELETVABQE'+intToStr(i),'') <> '' then
    	begin
			GrilleMontants.Cells[0,i] := 'n°'+intToStr(i);
			Q := OpenSQL('SELECT BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB FROM BANQUECP '
      						 +'WHERE BQ_GENERAL = "' + GetParamSocSecur('SO_TELETVABQE'+intToStr(i),'')
                               +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"', True); // 24/10/2006 YMO Multisociétés
			if not Q.eof then
      	begin
				GrilleMontants.Cells[1,i] := Q.FindField('BQ_ETABBQ').asString;
				GrilleMontants.Cells[2,i] := Q.FindField('BQ_GUICHET').asString;
				GrilleMontants.Cells[3,i] := Q.FindField('BQ_NUMEROCOMPTE').asString
        													 + Q.FindField('BQ_CLERIB').asString;
        end;
			GrilleMontants.Cells[4,i] := FormatFloat(GrilleMontants.ColFormats[4],0);
			// Génération de la référence
			GrilleMontants.Cells[5,i] := 'TVA'
      															+	FormatDateTime('YYYY',StrToDate(GetControlText('PERIODE')))		// Année en cours
                                    +	GetCodePeriodicite					// Périodicité
                                    +	FormatDateTime('MM',StrToDate(GetControlText('PERIODE_')))	// Mois de fin
                                    +	'0' + IntToStr(i)						// index (01,02,03)
                                    +	GetControlText('SIRET')			// siret
                                    +	'0000';											// '0000'
      Ferme(Q);
      end;

  // Début de la saisie dans la grille
  GrilleMontants.row := 1;
  GrilleMontants.col := 4;
  SetFocusControl('GRILLEMONTANTS');
end ;

procedure TOF_CPTVAEDI.OnArgument (S : String ) ;
begin
  Inherited ;
	// Paramètres Ecran
  GrilleMontants := THGrid(GetControl('GRILLEMONTANTS'));
	GrilleMontants.OnCellExit := OnCellExit;
	Ecran.OnKeyDown := FormKeyDown;

	// Paramétrage de la grille de saisie des montants
  GrilleMontants.ColEditables[0] := False; // 4 premières col non editables
  GrilleMontants.ColAligns[0] 	 := TaCenter;
  GrilleMontants.ColEditables[1] := False;
  GrilleMontants.ColAligns[1] 	 := TaCenter;
  GrilleMontants.ColEditables[2] := False;
  GrilleMontants.ColAligns[2] 	 := TaCenter;
  GrilleMontants.ColEditables[3] := False;
  GrilleMontants.ColAligns[3] 	 := TaCenter;
  GrilleMontants.ColTypes[4] 		 := 'F';
  GrilleMontants.ColFormats[4] 	 := '#,##0.00';
  GrilleMontants.ColAligns[4] 	 := TaRightJustify;
	GrilleMontants.ColLengths[5] 	 := 35; // Ref 35 car maxi. ??

	// Largeur des colonnes
  GrilleMontants.ColWidths[0] := 40;
  GrilleMontants.ColWidths[1] := 70;
  GrilleMontants.ColWidths[2] := 70;
  GrilleMontants.ColWidths[3] := 120;
  GrilleMontants.ColWidths[4] := 100;
  GrilleMontants.ColWidths[5] := 440;

	// Accès aux contrôles
  AccesControles;

end ;

procedure TOF_CPTVAEDI.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var	OkG , Vide : boolean;
begin
	inherited;
	if Not GrilleMontants.SynEnabled then
  	BEGIN
    	Key:=0;
      Exit;
    END;
	OkG:=(Screen.ActiveControl=GrilleMontants);
  Vide:=(Shift=[]);
	// Validation
  if key = VK_RETURN then
		if ((OkG) and (Vide)) then KEY:=VK_TAB;
end;

function TOF_CPTVAEDI.CoherenceMontants: Boolean;
var montant1, montant2, montant3, montantTotal : Double;
begin
	Result := True;
	Montant1 := Valeur(GrilleMontants.cells[4,1]);
	Montant2 := Valeur(GrilleMontants.cells[4,2]);
	Montant3 := Valeur(GrilleMontants.cells[4,3]);
  MontantTotal := Valeur(GetControlText('MONTANTTOTAL'));
  Result := (Montant1 + Montant2 + Montant3) = MontantTotal;
end;

procedure TOF_CPTVAEDI.OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var CumulMontant : Double;
begin
	if ACol = 4 then
  	begin
      if Trim(GrilleMontants.Cells[4, ARow]) = '' then
	      begin
        GrilleMontants.Cells[ACol, ARow] := FormatFloat(GrilleMontants.ColFormats[4],0);
        Cancel := True;
        Exit;
  	    end;

      if not IsNumeric(GrilleMontants.Cells[4, ARow]) then
    	  begin
        Cancel := True;
        Exit;
      	end
      else
      	begin   // Mise en place du cumul
        CumulMontant := Valeur(GrilleMontants.Cells[4, 1])
        							+ Valeur(GrilleMontants.Cells[4, 2])
                      + Valeur(GrilleMontants.Cells[4, 3]);
				SetControlText('CUMULGRILLE',FloatToStr(CumulMontant));
        end;

      GrilleMontants.Cells[ACol, ARow] := FormatFloat(GrilleMontants.ColFormats[4],Valeur(GrilleMontants.Cells[4, ARow]));
    end;
end;

function TOF_CPTVAEDI.CreerFichierEdiTVA: Boolean;
var
  QDos		       	      			: TQuery;
  TOB_TEMPO, TOB_DETAIL 			: TOB;
  TableValRub, PrefValRub 		: String;
 	Command, Buffer, sSQL				: String;
  CheminSource, RepSociete    : String;
  i                   				: Integer;
  F			    	       	  	 		: TextFile;
  DateDebut, DateFin  				: TDateTime;
begin
  Result := TRUE;
	// Récup des dates :
  DateDebut := StrToDate(GetControlText('PERIODE'));
  DateFin := StrToDate(GetControlText('PERIODE_'));
  TableValRub := LiasseEnv.FICHIERVALRUB;
  PrefValRub := LiasseEnv.FVAL;

	// Test existence application publifi.exe
  if not FileExists(V_PGI_ENV.RepLocal+ '\APP\tdi_publifi.exe') then
	  begin
    PGIINFO (TraduireMemoire(MessageListe[15]) ,Ecran.Caption);
    RESULT := FALSE;
    exit;
  	end;

	// Chargement des données de rubriques
  sSQL := 'SELECT * FROM ' + TableValRub
				+ ' WHERE ' + PrefValRub + '_MILLESIME = "' + LiasseEnv.MILLESIME + '"'
        +   ' AND ' + PrefValRub + '_DATEEFFET = "' + USDATETIME(LiasseEnv.DateEncours) + '"';
  QDos := OpenSQL(sSQL,true);
  TOB_TEMPO := TOB.Create('',Nil,-1);
  TOB_TEMPO.LoadDetailDB(TableValRub,'','',QDos,true,true);
  Ferme(QDOS);
  if TOB_TEMPO.Detail.count = 0 then
	  begin
    RESULT := FALSE;
    PGIINFO (TraduireMemoire(MessageListe[14]) ,Ecran.caption);
    exit;
  	end;

  if ctxPCL in V_PGI.PGIContexte
  	then CheminSource := V_PGI_ENV.pathDos + '\Tva.txt'
	  else CheminSource := V_PGI_ENV.RepLocal + '\STD\Tva.txt';

  DeleteFile(Pchar(CheminSource));

	try
    AssignFile(F, CheminSource);	{ Ouverture du fichier de sortie }

    if FileExists (CheminSource) then  Append (F) else Rewrite (F);

    Buffer := 'DATE' + #9 + FormatDateTime('yyyymmdd',DateDebut)
    								 + #9 + FormatDateTime('yyyymmdd',DateFin);
    Writeln(F,Buffer);
    Buffer := '00011' + #9 + FormatDateTime('yyyymmdd',DateDebut);
    Writeln(F,Buffer);
    Buffer := '00012' + #9 + FormatDateTime('yyyymmdd',DATEFIN);
    Writeln(F,Buffer);
    Buffer := '00014' + #9 + V_PGI_ENV.NoDossier;
    Writeln(F,Buffer);
    Buffer := '00015' + #9 + GetControlText('SIRETDECLA');
    Writeln(F,Buffer);
    Buffer := '00016' + #9 + GetControlText('ENTITEDECLA');
    Writeln(F,Buffer);
		if TCheckBox(GetControl('CESSATION')).checked  // Ajout 0017 : cessation d'activité
    	then Buffer := '00017' + #9 + 'OUI'
			else Buffer := '00017' + #9 + 'NON';
    Writeln(F,Buffer);

    // PARCOURS DE TOUTS LES RUB DE LA TOB
    //SI DE TYPE 0 (SAISIE) OU 2 (BD MODIFIE) ON ALIMENTE LE DICTIONAIRE DE DONNEE
    for i:= 0 to TOB_TEMPO.Detail.Count -1 do
	    begin
      TOB_DETAIL := TOB_TEMPO.Detail[i];
      if TOB_DETAIL.GetValue(PrefValRub + '_NUMRUB')= '00950' then EcritureRib (F);
      if TOB_DETAIL.GetValue(PrefValRub + '_NUMRUB')= '00100' then
        if ctxPCL in V_PGI.PGIContexte then
          begin
          Buffer := '00018' + #9 + 'CAB';
          Writeln(F,Buffer);
          end
        else
          begin
          Buffer := '00018' + #9 + 'ENT';
          Writeln(F,Buffer);
          end;
    	if TOB_DETAIL.GetValue(PrefValRub + '_NUMRUB')= '00400' then
		    begin
        Buffer := '00209' + #9 + GetControlText('CODEACTIVITE');
        Writeln(F,Buffer);
        Buffer := '00210' + #9 + GetControlText('SIRET');
        Writeln(F,Buffer);
        Buffer := '00211' + #9 + GetControlText('FRP');
        Writeln(F,Buffer);
		    end;

			// Ecriture de ligne principal
      Buffer := TOB_DETAIL.GetValue(PrefValRub + '_NUMRUB') + #9 + TOB_DETAIL.GetValue(PrefValRub + '_VALRUB');
      Writeln(F,Buffer);

      if TOB_DETAIL.GetValue(PrefValRub + '_NUMRUB') = '00113' then
	      begin
        Buffer := '00119' + #9 + GetParamSocSecur('SO_CODEPOSTAL','');
        Writeln(F,Buffer);
        Buffer := '00120' + #9 + GetParamSocSecur('SO_VILLE','');
        Writeln(F,Buffer);
  	    end ;
      if TOB_DETAIL.GetValue(PrefValRub + '_NUMRUB') = '00116' then
        if TOB_DETAIL.GetValue(PrefValRub + '_VALRUB') <> '' then
	        begin
          Buffer := '00117' + #9 + Copy(TOB_DETAIL.GetValue(PrefValRub + '_VALRUB'), 0, 5);
          Writeln(F,Buffer);
          Buffer := '00118' + #9 + Copy(TOB_DETAIL.GetValue(PrefValRub + '_VALRUB'), 6, length (TOB_DETAIL.GetValue(PrefValRub + '_VALRUB'))-5);
          Writeln(F,Buffer);
  	      end;
	  	end;

	finally
  	CloseFile(F);
  end;							// ---> fin du try

  TOB_TEMPO.Free;

  if ctxPCL in V_PGI.PGIContexte then
  	begin
    RepSociete := V_PGI_ENV.Com.dir + '\';
    CheminSource := V_PGI_ENV.pathDos + '\';
    end
  else
  	begin
    RepSociete := V_PGI_ENV.RepLocal + '\STD\'; // Avant : RepSociete = 'PME' ??
    CheminSource := V_PGI_ENV.RepLocal + '\STD\';
    end;

  Command := 		'TVA591,'												// Arg 1 : code pour nouveau comportement
  							+	'CA3,P,S,'                  	// Arg 2 à 4 :
  							+ V_PGI_ENV.NoDossier + ','	 		// Arg 5 : numéro de dossier
                + 'TVA.txt,'                  	// Arg 6 : nom du fichier généré
  							+ CheminSource + ','          	// Arg 7 : chemin du fichier généré
                + RepSociete + ','            	// Arg 8 : répertoire dossier
                + V_PGI_ENV.RepLocal + '\STD\';	// Arg 9 : répertoire des standards

  ShellExecute(0, PCHAR('open'),PCHAR('tdi_publifi.exe'),PCHAR(Command),Nil,SW_RESTORE);
END;

procedure TOF_CPTVAEDI.EcritureRib(var F: TextFile);
var
	Buffer : string;
  i : Integer;
begin
	// Mise en place des ribs concernant les 3 banques renseignées dans les paramètres société
	for i := 1 to 3 do
  begin
    if  GrilleMontants.cells[1,i] <> '' then
      begin
      Buffer := '00' + IntToStr(500+(i*10)) + #9 + GrilleMontants.cells[1,i]; 	 // Banque
      Writeln(F,Buffer);
			if IsNumeric(GrilleMontants.cells[4,i])			// Cas du montant
      	then Buffer := '00' + IntToStr(500+(i*10)+1) + #9 + FormatFloat('0.00',Valeur(GrilleMontants.cells[4,i]))
      	else Buffer := '00' + IntToStr(500+(i*10)+1) + #9 + '0,00';
      Writeln(F,Buffer);
      Buffer := '00' + IntToStr(500+(i*10)+2) + #9 + GrilleMontants.cells[5,i];  // Référence
      Writeln(F,Buffer);
      Buffer := '00' + IntToStr(500+(i*10)+3) + #9 + GrilleMontants.cells[2,i];  // Guichet
      Writeln(F,Buffer);
      Buffer := '00' + IntToStr(500+(i*10)+4) + #9 + GrilleMontants.cells[3,i];	 // Compte + Clé
      Writeln(F,Buffer);
      end;
  end;
end;

procedure TOF_CPTVAEDI.AccesControles;
begin
  SetControlEnabled('PERIODE',False);
  SetControlEnabled('PERIODE_',False);

  SetControlEnabled('RAISONSOCIALE',False);
  SetControlEnabled('ADRESSE',False);
  SetControlEnabled('CODEPOSTAL',False);
  SetControlEnabled('VILLE',False);

  SetControlEnabled('SIRET',False);
  SetControlEnabled('FRP',False);
  SetControlEnabled('CODEACTIVITE',False);
  SetControlEnabled('TVA',False);

// SetControlEnabled('ENTITEDECLA',False);    // A terme, renseigné auto
// SetControlEnabled('SIRETDECLA',False);     // A terme, renseigné auto

  SetControlEnabled('MONTANTTOTAL',False);
  SetControlEnabled('CUMULGRILLE',False);

end;

function TOF_CPTVAEDI.GetCodePeriodicite : String;
var Q : TQuery;
begin
	Result := '?';
  // Init des champs avec CA3_PARAMETRES
	Q := OpenSQL('SELECT CP3_PERIODICITE FROM CA3_PARAMETRE',True);
  if not Q.Eof then
	  begin
		Result := Q.FindField('CP3_PERIODICITE').asString;
		if (Result = 'MEN') or (Result = 'TRI') or (Result = 'ANN')
    	then	Result := Copy(Result,1,1)
      else 	Result := 'S';
  	end;
  Ferme(Q);
end;

Initialization
  registerclasses ( [ TOF_CPTVAEDI ] ) ;
end.
