unit InitFichCa3;

interface                                   
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97,HEnt1,Ed_Tools, HRotLab, hmsgbox, HFLabel, Courrier, dbtables,ent1,
  MenuOLG,Fe_MAin,PGIENV,HCtrls;

procedure ConnectionDico;
procedure CloseDico;
//procedure ImportBOBCa3;

implementation
// ajout me CA3
uses    LIA_Edition{,AideLiasse}, LIA_commun, LIA_FCT_EXTERNE,UTOMPARAMCA3;


PROCEDURE RenseigneDateCa3;
var
Q1        : TQuery;
begin
LiasseEnv.DateEncours := iDate1900;
if (V_PGI_Env <>Nil)  and (V_PGI_Env.ModeFonc='MULTI') then
   Q1 := OpenSql ('SELECT CP3_DATEENCOURS FROM CA3_PARAMETRE WHERE CP3_CODEPER='+ IntToSTr (V_PGI_ENV.Codeper), TRUE)
else
   Q1 := OpenSql ('SELECT CP3_DATEENCOURS FROM CA3_PARAMETRE WHERE CP3_CODEPER=1', TRUE);
   if not Q1.EOF then
   begin
     LiasseEnv.DateEncours := Q1.Fields[0].Asdatetime;
     if LiasseEnv.DateEncours = 0 then
     begin
       LiasseEnv.DateEncours := VH^.Encours.Deb;
       if (LiasseEnv.dateEncours <> VH^.CPExoRef.Deb) then
       LiasseEnv.DateEncours := VH^.CPExoRef.Deb;
     end;
   end else
   begin
        LiasseEnv.DateEncours := VH^.Encours.Deb;
        if (LiasseEnv.dateEncours <> VH^.CPExoRef.Deb) then
        LiasseEnv.DateEncours := VH^.CPExoRef.Deb;
   end;
if LiasseEnv.dateEncours < VH^.CPExoRef.Deb then
LiasseEnv.DateEncours := VH^.CPExoRef.Deb;
Ferme (Q1);
end;

procedure ConnectionDico;
begin
          VLIASSE^.RepDeclaration :='3310';
          VLIASSE^.NomExe :='CA32002s5';

          LiasseEnv := TLiasseEnv.Create;
             //LECTURE DU FICHIER .INI
          IF NOT fct_VERIF_FICHIER_INI(VLIASSE^.RepDeclaration, VLIASSE^.NomExe,TRUE) THEN
          BEGIN
               FMENUG.ForceClose := TRUE;
               FMENUG.Close;
               exit;
          END;
          RenseigneDateCa3; // ajout me
{
          IF NOT fct_IMPORT_DONNEES_TABLE(Liasse_RepDeclaration) THEN
          BEGIN
               PGIINFO('PROBLEME SUR fct_IMPORT_DONNEES_TABLE','TRAITEMENT IMPOSSIBLE');
               FMENUG.ForceClose := TRUE;
               FMENUG.Close;
               exit;
          END;
}
// ajout me pour ca3
          //INITIALISATION DU RUBDICO
          IF NOT fct_OUV_RUBDICO(VLIASSE^.RepDeclaration+ '\'+LiasseEnv.MILLESIME) then
          BEGIN
               FMENUG.ForceClose := TRUE;
               FMENUG.Close;
               exit;
          END;
          //FONCTION SPECIFIQUE DU MOTEUR
          vb_SetUserInfoBd (InfoBdPgi);

          //MAJ DU CODEPER DANS LE DICO
           IF NOT fct_MAJ_CODEPER_DICO(V_PGI.CodeSociete) THEN
          BEGIN
               FMENUG.ForceClose := TRUE;
               FMENUG.Close;
               exit;
          END;
          if not ExisteSQL ('SELECT * FROM F3310_VALRUB_STRUC') then
          fct_IMPORT_DONNEES_BOB;
          if (V_PGI_Env <>Nil)  and (V_PGI_Env.ModeFonc='MULTI') then
          begin
                    // fct_IMPORT_DONNEES_BOB(TRUE,TRUE);
                    // ajout me 01/07/2002 MajInfoparamCA3;
          end;

          //AFFICHAGE DU LIBELLE AVEC OU SANS COMPTA
          if fct_VALRUB_VIDE then
          begin

               FMenuG.OutLook.ActiveGroup:=1;
               // ajout me  01/07/2002
               //    AGlLanceFiche ('CP','FICHEPARAMCA3',LiasseEnv.Codeper,LiasseEnv.Codeper,'');
               exit;
          end
          else
          begin
              if  (NombreENREG_VALRUB <> NombreENREG_VALRUBSTRUC)
              and (NombreENREG_VALRUB > 0) then
              begin
                   PGIInfo('diff','PREMIERE UTILISATION');
              end;
              //ALIMENTATION DU RUBDICO A PARTIR DU VALRUB
              IF NOT fct_INIT_RUBDICO(true) THEN
              BEGIN
                   FMENUG.ForceClose := TRUE;
                   FMENUG.Close;
                   exit;
              END;
              //ALIMENTATION DU RUBDICO A PARTIR DU VALRUB
              IF NOT fct_INIT_RUBDICO(true) THEN
              BEGIN
                   FMENUG.ForceClose := TRUE;
                   FMENUG.Close;
                   exit;
              END;
              //AFFICHAGE DU LIBELLE AVEC OU SANS COMPTA
              IF NOT fct_AFFICHAGE_EXERCICE THEN
              BEGIN
                   FMENUG.ForceClose := TRUE;
                   FMENUG.Close;
                   exit;
             END;
          END;
          V_PGI.QRPdf := TRUE;
          //Pour ne pas avoir un panel vide
        //  FMenuG.ChoixModule;
end;

procedure CloseDico;
begin
  if LiasseEnv <> nil then
  begin
    LiasseEnv.Free;
    vb_FermeDictionnaire;
  end;
end;
{
procedure ImportBOBCa3;
begin
//               fct_IMPORT_DONNEES_BOB(Liasse_RepDeclaration);
          if (V_PGI_Env <>Nil)  and (V_PGI_Env.ModeFonc='MULTI') then
                    fct_IMPORT_DONNEES_BOB(TRUE,TRUE)
          else
                    fct_IMPORT_DONNEES_BOB(FALSE,TRUE);

end;
}
end.
