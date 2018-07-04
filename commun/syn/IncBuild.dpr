(**************************** PROGRAMME *********************************
Auteur  ...... : Laurent Abélard
Créé le ...... : 20/11/2007
Modifié le ... : 23/01/2008
Description .. : Incrémentation du Numéro de Build d'un fichier de
Suite ........ : ressource .res
Suite ........ : IncBuild.exe NomFichier.res
Mots clefs ... : Ressource, .res, .rc, Build
************************************************************************)
Program IncBuild;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  Windows,
  Classes,
  StrUtils,
  IncBuildSources in '..\LIB\IncBuildSources.pas';

(******************************************************************
  Conversion d'une chaine ANSI dans le Code de Page OEM
******************************************************************)
Function StrAnsiToOEM( Const Texte : String ) : String;
Var
  TailleChaine : Integer;

Begin
  TailleChaine := Length(Texte);
  SetLength( Result, TailleChaine );
  AnsiToOemBuff( @Texte[1], @Result[1], TailleChaine );
End;


(******************************************************************
  Procedure de sortie standard dans le Code Page OEM
******************************************************************)
Procedure Afficher( Const Texte : String );
Begin
  Writeln( 'IncBuild : ' + StrAnsiToOEM( Texte ) );
End;

(******************************************************************
  Point d'entrée du programme principal
******************************************************************)
Var
  TypeValeurForcee : BUILD_OU_VERSION;
  TypeParametre    : String;
  ValeurForcee     : String;
  NomFichier       : String;
  Parametre        : String;
  CheminXml        : String;
  Resultat         : String;
  VersionProduit   : String;
  i                : Cardinal;

Begin

  ExitCode := 0;

  NomFichier := ParamStr(1);
  If NomFichier = '' Then
  Begin
    Afficher( 'Usage : ' );
    Afficher( ExtractFileName( ParamStr(0) ) + ' Nom_de_fichier.res [-x:Chemin_fichier_log.xml] '
                                             + '[-v:Numéro_de_Version_Forcée | -b:Numéro_de_Build_Forcée]' );
    Afficher( 'Où le numéro de version doit avoir la forme Majeur.Mineure.Release.build' );
    Afficher( 'Retourne 1 en cas d''erreur sinon 0.' );
    Exit;
  End;

    //** Récupération des paramètres **//
  TypeValeurForcee := BUILD_FORCE;
  ValeurForcee     := '';
  CheminXml        := '';
  VersionProduit    :='';

  For i := 2 To ParamCount() Do
  Begin
    If Length(ParamStr(i)) > 3 Then
    Begin

      TypeParametre := LeftStr( ParamStr(i), 3 );
      Parametre     := AnsiMidStr( ParamStr(i), 4, 10000 );

        //** Paramètre du Chemin d'un Fichier de Résultat en XML **//
      If TypeParametre = '-x:' Then
      Begin
        CheminXml := Parametre;
      End

        //** Paramètre Numéro de Build Forcé **//
      Else If TypeParametre = '-b:' Then
      Begin

        If ValeurForcee <> '' Then
        Begin
          Afficher( 'Vous ne pouvez pas spécifier un Numéro de Build Forcé en même temps qu''un Numéro de Version Forcée.' );
          Exit;
        End;

        TypeValeurForcee := BUILD_FORCE;
        ValeurForcee     := Parametre;

      End

        //** Paramètre Numéro de Version Forcée **//
      Else If TypeParametre = '-v:' Then
      Begin

        If ValeurForcee <> '' Then
        Begin
          Afficher( 'Vous ne pouvez pas spécifier un Numéro de Version Forcée en même temps qu''un Numéro de Build Forcé.' );
          Exit;
        End;

        TypeValeurForcee := VERSION_FORCE;
        ValeurForcee     := Parametre;

      End
      Else If TypeParametre = '-p:' Then
      Begin

        VersionProduit     := Parametre;

      End;
    End;
  End;

    //** Traitement du Fichier .res **//
  Try

    Resultat := IncrementerBuild( NomFichier, CheminXml, ValeurForcee, TypeValeurForcee, VersionProduit );
    If Resultat <> '' Then
      Afficher( Resultat );

  Except

    On Ex : Exception Do
    Begin
      Afficher( Ex.Message );
      Afficher( 'Avec les paramètres : ' + NomFichier +' , '+ CheminXml +' , '+ ValeurForcee );
      ExitCode := 1;
    End;

  End;

End.
