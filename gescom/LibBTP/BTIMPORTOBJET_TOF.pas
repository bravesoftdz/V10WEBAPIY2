{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 17/06/2008
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYIMPORTOBJET ()
Mots clefs ... : TOF;YYIMPORTOBJET
*****************************************************************}
Unit BTIMPORTOBJET_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
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
     ubob,
     UTOF ; 

Type
  TOF_BTIMPORTOBJET = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;


 	Private
    //
    
    //
  end ;
  
Implementation

procedure TOF_BTIMPORTOBJET.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTOBJET.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTOBJET.OnUpdate ;
Var FileName : String;
    RepBob	 : String;
begin
  Inherited ;

  RepBob	 := ExtractFilePath(GetControlText('FICIMPORT'));
  FileName := ExtractFileName(GetControlText('FICIMPORT'));

  if FileName = '' then
  	 begin
     PGIBox('Veuillez sélectionner un fichier');
     SetFocusControl('FICIMPORT');
     exit;
     end;

  if PGIAsk('Confirmez-vous le traitement de ce fichier ?') <> mrYes then exit;

  AGLIntegreBob(RepBob + FileName);

  case AGLIntegreBob(RepBob + FileName, FALSE, TRUE) of
       0 : begin
           if V_PGI.SAV then Pgiinfo('Intégration de : '+ RepBob + FileName, TitreHalley);//Resultif not LIA_JOURNAL_EVENEMENT(sTempo) then Result := -1;
           //if copy(FileName,9,1) = 'M' then Result := 1; //SI BOB AVEC MENU, ON REND 1 POUR SORTIR DE L'APPLICATION
           end;
       1 : if V_PGI.SAV then Pgiinfo('Intégration déjà effectuée :'+ FileName, TitreHalley);// Intégration déjà effectuée
      -1 : begin // Erreur d'écriture dans la table YMYBOBS
           if V_PGI.SAV then PGIInfo('Erreur d''écriture dans la table YMYBOBS :'+ RepBob + FileName,'PCL_IMPORT_BOB');
           end;
      -2 : begin // Erreur d'intégration dans la fonction AglImportBob
           if V_PGI.SAV then PGIInfo('Erreur d''intégration dans la fonction AglImportBob :'+ RepBob + FileName,'PCL_IMPORT_BOB');
           end;
      -3 : begin //Erreur de lecture du fichier BOB.
           if V_PGI.SAV then PGIInfo('Erreur de lecture du fichier BOB :'+ RepBob + FileName,'PCL_IMPORT_BOB');
           end;
      -4 : begin // Erreur inconnue.
           if V_PGI.SAV then PGIInfo('Erreur inconnue :'+ RepBob + FileName,'PCL_IMPORT_BOB');
           end;
  end;

end ;

procedure TOF_BTIMPORTOBJET.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTOBJET.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTOBJET.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTOBJET.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTOBJET.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTIMPORTOBJET ] ) ; 
end.

