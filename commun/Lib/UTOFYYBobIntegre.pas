{***********UNITE*************************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BOBINTEGRE ()
Mots clefs ... : TOF;BOBINTEGRE
*****************************************************************}
Unit UTOFYYBobIntegre ;
/////////////////////////////////////////////////////////////////
Interface
/////////////////////////////////////////////////////////////////
Uses
   UTOF ;
/////////////////////////////////////////////////////////////////
Type
   TOF_YYBOBINTEGRE = Class (TOF)
         procedure OnArgument ( sListeParam_p : String ) ; override ;
      private
         procedure OnClick_BDELETE( Sender : TObject );
         procedure PrepareDelete;
   end ;
/////////////////////////////////////////////////////////////////
Implementation
/////////////////////////////////////////////////////////////////
Uses
   HTB97, HDB, hMsgBox, Mul, Controls, Classes, AGLInit,
   HEnt1, LicUtil, HCtrls, sysutils, hqry, HStatus;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 01/04/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_YYBOBINTEGRE.OnArgument (sListeParam_p : String ) ;
begin
   Inherited;
   Ecran.HelpContext := 2611300000;
   THDBGrid( GetControl( 'FLISTE' )).MultiSelection := V_PGI.Superviseur;//(V_PGI.PassWord = CryptageSt(DayPass(Date)));
	SetControlVisible('BDELETE', V_PGI.Superviseur);  //(V_PGI.PassWord = CryptageSt(DayPass(Date))));
   TToolbarButton97( GetControl( 'BDELETE' ) ).OnClick := OnClick_BDELETE;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/04/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_YYBOBINTEGRE.OnClick_BDELETE( Sender : TObject );
var
   sBobName_l : string;
begin
//   PrepareDelete;
   sBobName_l := TFMul(Ecran).Q.FindField('YB_BOBNAME').asString;
   if PGIAsk('Vous allez supprimer le BOB "'+ sBobName_l + '" : ' + #13#10 +
             ' les données qu''il contient pourront être réimportées' + #13#10 +
             ' et donc écraser les valeurs actuelles si sa version est ' + #13#10 +
             ' antérieur à la SocRef en cours' + #13#10 +
             'Confirmez-vous?',
             'Intégration de BOB') = mrYes then
   begin
      ExecuteSQL ('DELETE FROM YMYBOBS WHERE YB_BOBNAME = "' + sBobName_l + '"');
      AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/10/2003
Modifié le ... :   /  /
Description .. : Parcours la grille et effectue le traitement
Mots clefs ... :
*****************************************************************}
procedure TOF_YYBOBINTEGRE.PrepareDelete;
var
   dbgGrille_l : THDBGrid;
   qQuery_l : THQuery;
   nRow_l : integer;
   sDelete_l, sMessage_l : string;
begin
   dbgGrille_l := THDBGrid(GetControl('FLISTE'));
   qQuery_l := THQuery(TFMul(Ecran).Q);

   nRow_l := 0;
   if dbgGrille_l.AllSelected then
   begin
      // Toutes les lignes sélectionnées
      InitMove( qQuery_l.RecordCount, '' );
      qQuery_l.First;
      for nRow_l := 0 to qQuery_l.RecordCount - 1 do
      begin
         MoveCur(False);
         if sDelete_l <> '' then sDelete_l := sDelete_l + ', ';
         sDelete_l := sDelete_l + '"' + qQuery_l.FindField('YB_BOBNAME').asString + '"';
         sMessage_l := sMessage_l + '- ' + qQuery_l.FindField('YB_BOBNAME').asString +  #13#10;
         qQuery_l.Next;
      end;
      dbgGrille_l.AllSelected:=False;
      FiniMove;
   end
   else
   begin
      if dbgGrille_l.NbSelected <> 0 then
      begin
         // Quelques lignes sélectionnées
         InitMove( dbgGrille_l.NbSelected, '' );
         for nRow_l := 0 to dbgGrille_l.NbSelected - 1 do
         begin
            MoveCur(False);
            dbgGrille_l.GotoLeBookmark( nRow_l );
            {$IFDEF EAGLCLIENT}
            qQuery_l.TQ.Seek(dbgGrille_l.Row - 1) ;
            {$ENDIF}
            if sDelete_l <> '' then sDelete_l := sDelete_l + ', ';
            sDelete_l := sDelete_l + '"' + qQuery_l.FindField('YB_BOBNAME').asString + '"';
            sMessage_l := sMessage_l + '- ' + qQuery_l.FindField('YB_BOBNAME').asString +  #13#10;
         end;
         dbgGrille_l.ClearSelected;
         FiniMove;
      end;
   end;
   if ( sDelete_l = '' ) then
      PGIInfo( 'Aucuns élément sélectionné', 'Intégration de BOB' )
   else if ( PGIAsk('Vous allez supprimer les informations des BOB suivants :' + #13#10 +
               sMessage_l + 'Confirmez-vous?',
               'Intégration de BOB') = mrYes ) then
   begin
      ExecuteSQL ('DELETE FROM YMYBOBS WHERE YB_BOBNAME IN ( '+ sDelete_l + ')');
      AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
   end;
end;

/////////////////////////////////////////////////////////////////
Initialization
  registerclasses ( [ TOF_YYBOBINTEGRE ] ) ;
end.

