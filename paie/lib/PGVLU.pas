{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/09/2002
Modifié le ... : 20/01/2004
Description .. : Edition de la VLU
Suite ........ : concerne organisme Assedic, ducs dossier, paiement 
Suite ........ : groupé
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
{
 PT1 : 23/02/2004 : V_5.0 MF FQ 10648  : mise en place dans les états
                          chaînées des éditions  de la DUCS
 PT2 : MF 25/07/2005 : V_6.04 FQ  12445 : correction en tête VLU caisse
                     destinataire
 PT3 : 30/03/2007 : V_702 MF : Modification : l'édition de la VLU est réalisée
                              à partir de l'édition des DUCS (menu 42354), de
                              l'impression VLU (menu 42303), des éditions chaînées
                              et du process server
}
unit PGVLU;

interface
uses
    Classes,Comctrls, HCtrls,HEnt1 ,
    SysUtils, // PT3  pour déclaration DateToStr
    P5def,    // PT3  pour déclaration CreeJnalEvt
{$IFDEF EAGLCLIENT}
      UtileAGL;
//unused      MaineAgl;
{$ELSE}
      EdtREtat;
{$ENDIF}      

    Procedure EditVLU(Debper,FinPer : TDateTime; Organisme : String;Pages : TPageControl; Chaine : boolean; Apercu : boolean; var StSQL : string); // PT1 PT3

implementation

Procedure EditVLU(Debper,FinPer : TDateTime; Organisme : String;Pages : TPageControl; Chaine : boolean; Apercu : boolean; var StSQL : string);  // PT1 PT3
var
//PT3   StSQL                         : string; // est passé en paramètre variable
   Trace, TraceE                      : TStringList; //PT3 pour maj jnal des événements

begin
// d PT3
  // maj jnal des événements
  Trace := TStringList.Create;
  TraceE := TStringList.Create;
  Trace.Add ('DUCS : Traitement d''impression des VLU pour la période du '+ DateToStr(DebPer) + ' au '+ DateToStr(FinPer));
// f PT3

   StSQL := 'SELECT POG_ETABLISSEMENT,'+
            'POG_CAISSEDESTIN,'+ // PT2 FQ 12445
            'ET_LIBELLE,ET_SIRET,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,'+
            'ET_CODEPOSTAL,ET_VILLE,ET_APE,ET_TELEPHONE,ET_FAX,'+
            'POG_ORGANISME,POG_NUMINTERNE,POG_LIBELLE,'+
            'POG_ADRESSE1,POG_ADRESSE2,POG_ADRESSE3,POG_CODEPOSTAL,'+
            'POG_VILLE,'+
            'PDU_ETABLISSEMENT,PDU_ORGANISME,PDU_ABREGEPERIODE,'+
            'PDU_DATEDEBUT, PDU_DATEFIN '+
            'FROM ORGANISMEPAIE '+
            'LEFT JOIN DUCSENTETE ON POG_ORGANISME=PDU_ORGANISME '+
            'AND POG_ETABLISSEMENT = PDU_ETABLISSEMENT '+ //@@
            'LEFT JOIN ETABLISS on POG_ETABLISSEMENT = ET_ETABLISSEMENT '+
            'WHERE POG_ORGANISME ="002" AND POG_DUCSDOSSIER="X" AND '+
            'POG_PAIEGROUPE="X" AND PDU_DATEDEBUT="'+UsDateTime(Debper)+'"'+
            ' AND PDU_DATEFIN="'+USDAteTime(FinPer)+'" '+
            'ORDER BY POG_CAISSEDESTIN DESC, PDU_ETABLISSEMENT'; // PT2 FQ 12445
{d PT3   if (Chaine = True)  then
     begin
      LanceEtat('E','PDU','PVU',Apercu,False,False,Pages,StSQL,'',False);
      V_PGI.NoPrintDialog := True; // pas d'affichage de la fenêtre d'impression
     end
   else}
   if (Chaine <> True)  then
   // si Etat chaîné ou Impression VLU (menu 42303)
     LanceEtat('E','PDU','PVU',Apercu,False,False,Pages,StSQL,'',False);

  // maj jnal des événements
   Trace.Add ('Fin du traitement d''impression des VLU ');
   CreeJnalEvt('001', '022', 'OK', nil, nil, Trace, TraceE);

   Trace.Free;
   TraceE.Free;
// f PT3
end;

end.
