{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/05/2005
Modifié le ... :   /  /
Description .. : Gestion de la fenetre de saisie d'un motif de suppression
                 pour un appel
Mots clefs ... : TOF;UTOFANNULAPP
*****************************************************************}
Unit UTofAnnulApp ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,
{$else}
     eMul,
     Maineagl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Vierge,
 		 AglInit,
     HRichOle,
     UTOF ;

Type
  TOF_ANNULAPP = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    CodeAffaire : String;
    Etat        : String;
    P0          : String;
    P1          : String;
    P2          : String;
    P3          : String;
    Avenant     : String;

    DateSup     : TdateTime;
    UneTob : TOB;
    procedure ControleChamp(Champ, Valeur: String);

  end ;

Implementation

procedure TOF_ANNULAPP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_ANNULAPP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_ANNULAPP.OnUpdate ;
Var TobOLE : Tob;
    Motif	 : THRichEditOLE;
begin
  Inherited ;

  Motif := THRichEditOLE(GetControl('LO_MOTIF'));

  if (Length(Motif.Text) = 0) or (Motif.Text = #$D#$A)  then
     Begin
     if Etat = 'ANN' then
	      PGIBox(TraduireMemoire('Le motif de suppression est obligatoire'),TraduireMemoire('Motif obligatoire'))
   	 else if Etat = 'REA' then
	      PGIBox(TraduireMemoire('Le motif de l''appel est obligatoire'),TraduireMemoire('Motif obligatoire'));
     exit;
     end;

  //Chargement des Tob liens OLE
  TobOle := Tob.Create('LIENSOLE' ,Nil, -1);

  TobOle.PutValue('LO_TABLEBLOB', 'APP');
  TobOle.PutValue('LO_QUALIFIANTBLOB', 'MOT');
  TobOle.PutValue('LO_EMPLOIBLOB', Etat);
  TobOle.PutValue('LO_IDENTIFIANT', CodeAffaire);
  TobOle.PutValue('LO_RANGBLOB', 1);

  if Etat = 'ANN' then
	   TobOle.PutValue('LO_LIBELLE', 'Annulation Affaire ' + CodeAffaire + 'le ' + DateToStr(Datesup))
  Else if Etat = 'REA' then
	   TobOle.PutValue('LO_LIBELLE', 'Intervention Téléphonique sur ' + CodeAffaire + 'le ' + DateToStr(Datesup));

  TobOle.PutValue('LO_PRIVE', '-');
  TobOle.PutValue('LO_DATEBLOB', DateSup);
  TobOle.PutValue('LO_OBJET', GetControlText('LO_MOTIF'));

  TobOle.InsertOrUpdateDB(false);
  TobOle.free;
  LaTob.PutValue('RETOUR','X');

end ;

procedure TOF_ANNULAPP.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_ANNULAPP.OnArgument (S : String ) ;
var Critere : String;
    Valeur  : String;
    Champ   : String;
    X       : integer;
begin
  Inherited ;

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
     Critere := (Trim(ReadTokenSt(S)));
    end;

    setControlText('AFF_AFFAIRE', 'APPEL N° ' + P1+P2+P3);

    If Etat = 'ANN' then
       Begin
		  Ecran.Caption := 'Saisir le Motif de l''annulation d''un Appel';
	     SetControlText('TAFF_MOTIF', 'MOTIF D''ANNULATION');
	     SetControlText('TAFF_DATE', 'Annulé le ' + DateToStr(DateSup));
       End
    Else if Etat = 'REA' Then
       Begin
       Ecran.Caption := 'Saisir le Descriptif de l''intervention';
	    SetControlText('TAFF_MOTIF', 'DEPANNAGE');
       SetControlText('TAFF_DATE', 'Réalisé le ' + DateToStr(DateSup));
       end;

    OnDisplay;

end ;

Procedure TOF_ANNULAPP.ControleChamp(Champ : String;Valeur : String);
Begin

  if      Champ = 'CREATION' Then TFVierge(ecran).TypeAction := TaCreat
  Else if Champ = 'CONSULTE' Then TFVierge(ecran).TypeAction := TaConsult
  Else if Champ = 'APPEL'    Then CodeAffaire := Valeur
  Else if Champ = 'AFFAIRE0' Then P0 := Valeur
  Else if Champ = 'AFFAIRE1' Then P1 := Valeur
  Else if Champ = 'AFFAIRE2' Then P2 := Valeur
  Else if Champ = 'AFFAIRE3' Then P3 := Valeur
  Else if Champ = 'AVENANT'  Then Avenant := Valeur
  Else if Champ = 'DATE'     Then DateSup := StrToDate(Valeur)
  Else If Champ = 'ETAT'     Then Etat := Valeur;

end;

procedure TOF_ANNULAPP.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_ANNULAPP.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_ANNULAPP.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_ANNULAPP ] ) ;
end.

