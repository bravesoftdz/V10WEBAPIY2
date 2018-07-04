{***********UNITE*************************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 27/06/2003
Modifié le ... : 27/06/2003
Description .. : Moulinette pour la fusion des champs date et des champs 
Suite ........ : heure de la table juevenement :
Suite ........ : JEV_DATE et JEV_HEUREDEB
Suite ........ : JEV_DATEFIN et JEV_HEUREFIN
Suite ........ : JEV_ALERTEDATE et JEV_ALERTEHEURE
Suite ........ : avant suppresion des champs JEV_HEUREDEB, 
Suite ........ : JEV_HEUREFIN, JEV_ALERTEHEURE
Mots clefs ... :
*****************************************************************}
unit UConvertirDateHeureEvenement;
/////////////////////////////////////////////////////////////////
interface
/////////////////////////////////////////////////////////////////
uses
{$IFNDEF EAGLCLIENT}
   dbTables,
{$ELSE}
{$ENDIF}
UTOB, Controls;
/////////////////////////////////////////////////////////////////
   procedure ConvertirDateHeureEvenement;
   procedure ConvertirLesDatesHeures( OBEvenement_p : TOB );
   procedure ConvertirUneDateHeure( dDate_p : TDate; tHeure_p : TTime; var dtDateHeure_p : TDateTime );

/////////////////////////////////////////////////////////////////
implementation
/////////////////////////////////////////////////////////////////
uses
   HEnt1, sysutils, HCtrls;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 20/05/2003
Modifié le ... :
Description .. : Après création du champ ANL_CONVRANG : integer
Mots clefs ... :
*****************************************************************}
Procedure ConvertirDateHeureEvenement;
var
   OBEvenement_l : TOB;
   QRYEvenement_l : TQuery;
   nEvtInd_l, nEvtNb_l : integer;
begin
   OBEvenement_l := TOB.Create( 'JUEVENEMENT', nil, -1 );
   QRYEvenement_l := OpenSQL( 'select JEV_GUIDEVT, JEV_DATE, JEV_HEUREDEB, ' +
                              '       JEV_DATEFIN, JEV_HEUREFIN, ' +
                              '       JEV_ALERTEDATE, JEV_ALERTEHEURE, JEV_DUREE ' +
                              'from JUEVENEMENT ',
                              true );

   OBEvenement_l.LoadDetailDB('JUEVENEMENT', '', '', QRYEvenement_l, false );
   nEvtNb_l := OBEvenement_l.Detail.Count;
   For nEvtInd_l := 0 to nEvtNb_l - 1 do
   begin
      ConvertirLesDatesHeures( OBEvenement_l.Detail[nEvtInd_l] );
      OBEvenement_l.Detail[nEvtInd_l].UpdateDB;
   end;
   Ferme(QRYEvenement_l);
   OBEvenement_l.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 20/05/2003
Modifié le ... :   /  /    
Description .. : Traitement sur les champs ANL_CONVLIB, ANL_CONVSUITE,
Suite ........ : ANL_REFLIEN et ANL_CONVTXT
Mots clefs ... :
*****************************************************************}
procedure ConvertirLesDatesHeures( OBEvenement_p : TOB );
var
   dtDateHeureDeb_l, dtDateHeureFin_l, dtDateHeureAlerte_l : TDateTime;
   dDateDeb_l, dDateFin_l, dDateAlerte_l : TDate;
   tHeureDeb_l, tHeureFin_l, tHeureAlerte_l : TTime;
begin
   dDateDeb_l := OBEvenement_p.GetValue('JEV_DATE');
   tHeureDeb_l := OBEvenement_p.GetValue('JEV_HEUREDEB');
   ConvertirUneDateHeure( dDateDeb_l, tHeureDeb_l, dtDateHeureDeb_l );

   dDateFin_l := OBEvenement_p.GetValue('JEV_DATEFIN');
   tHeureFin_l := OBEvenement_p.GetValue('JEV_HEUREFIN');
   ConvertirUneDateHeure( dDateFin_l, tHeureFin_l, dtDateHeureFin_l );

   dDateAlerte_l := OBEvenement_p.GetValue('JEV_ALERTEDATE');
   tHeureAlerte_l := OBEvenement_p.GetValue('JEV_ALERTEHEURE');
   ConvertirUneDateHeure( dDateAlerte_l, tHeureAlerte_l, dtDateHeureAlerte_l );

   OBEvenement_p.PutValue('JEV_DATE', dtDateHeureDeb_l );
   OBEvenement_p.PutValue('JEV_DATEFIN', dtDateHeureFin_l );
   OBEvenement_p.PutValue('JEV_ALERTEDATE', dtDateHeureAlerte_l );
   OBEvenement_p.PutValue('JEV_DUREE', dtDateHeureFin_l - dtDateHeureDeb_l );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 27/06/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure ConvertirUneDateHeure( dDate_p : TDate; tHeure_p : TTime; var dtDateHeure_p : TDateTime );
begin
   // Date 01/01/1900
   if ( dDate_p = iDate1900 ) then
      dtDateHeure_p := dDate_p
   // Date et Heure 02/06/2003 10:20:45 : Déjà mis à jour
   else if ( dDate_p <> iDate1900 ) and ( Length( DateTimeToStr( dDate_p ) ) = 19 ) then
      dtDateHeure_p := dDate_p
   // Date 02/06/2003 et Heure 01/01/1900
   else if ( dDate_p <> iDate1900 ) and ( tHeure_p = iDate1900) then
      dtDateHeure_p := dDate_p
   // Date 02/06/2003 et Heure 30/12/1899 10:20:45
   else if ( dDate_p <> iDate1900 ) and ( tHeure_p < iDate1900) then
      dtDateHeure_p := dDate_p + tHeure_p
   // Date 02/06/2003 et Heure 02/06/2003 : cas en erreur  
   else
      dtDateHeure_p := dDate_p;
end;
/////////////////////////////////////////////////////////////////
end.
