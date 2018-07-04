{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/02/2002
Modifié le ... : 14/02/2002
Description .. : Source TOF de la FICHE : FACT_ECLAT ()
Suite ........ : ecran qui permet de changer l'optiond e fatcuration éclatée 
Suite ........ : par assistant et qui fait les traitement voulu dans les tables 
Suite ........ : en fct du nouveau choix fait
Mots clefs ... : TOF;UTOFFACT_ECLAT
*****************************************************************}
Unit UTOFFACT_ECLAT ;

Interface

Uses StdCtrls,  Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db, 
     dbtables, FE_Main,
{$ELSE}
    Maineagl,
{$ENDIF}
     forms,ParamSoc,DicoAf ,sysutils,UtilSocAf,FactTob,
     ComCtrls,HCtrls, HEnt1, HMsgBox,Ent1,
     UTOF,AffaireUtil,EntGc,Utob ,FactUtil,FactComm; 

Type
  TOF_FACT_ECLAT = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnUpdate                 ; override ;
    Procedure BoucleEclatFact;
  Private
    AfFactParRes : string;
    Erreur:Boolean;
    end ;
Procedure AFLanceFiche_FAct_Eclat;

Implementation


procedure TOF_FACT_ECLAT.OnArgument (S : String ) ;
begin
  Inherited ;
Erreur:=False;
AfFactParRes := GetParamSoc('So_AfFactParres');
SetControltext ('AFFACTPARRES',AfFactParRes);
end ;

procedure TOF_FACT_ECLAT.OnClose ;

begin
  Inherited ;
If Erreur then begin
   LastError:=(-1);
   LastErrorMsg:='';
   Erreur:=false;
   exit;     // il y eu une erreur, on reste sur la même fiche
   end;
end ;

procedure TOF_FACT_ECLAT.OnUpdate;
var text,titre:string;
begin
  Inherited ;
titre:='Changement facture éclatée';
If (GetControlText('AFFACTPARRES1') = '' ) or
   (GetControlText('AFFACTPARRES1') = AfFactParRes ) then begin
    PGIInfo('Vous n''avez pas changé de valeur ou valeur vide',titre);
    Lasterror:=1;
    Erreur:=True;
    exit;
    end
else begin
     if (GetControlText('AFFACTPARRES1') = 'SAN') then Text :='Confirmez-vous la suppression de la facturation éclatée par ressource'
     else text:='Confirmez-vous la mise en place de la facturation éclatée par ressource avec alimentation des tables concernées,après avoir mis à jour vos paramètres';
     If PGIAsk(traduitGA(text),titre)<> mrYes then exit;
     SetParamSOc('SO_AFFACTPARRES',GetControlText('AFFACTPARRES1')); // stocke nouvelle valeur
     if (GetControlText('AFFACTPARRES1') <> 'SAN') then begin  // il faut saisir les autres paramètres
          ParamSociete(False,'','SCO_FACTUREECLATEE','',nil,ChargePageSocGA,SauvePageSocGA,InterfaceSocGA,110000220);
          ChargeSocieteHalley;
          end;
     if (GetControlText('AFFACTPARRES1') <> 'SAN') and
       ((getparamsoc('SO_AFFACTMODEPEC')='')
        or (GetparamSoc('SO_AFFACTPRESDEFAUT')='')
        or (GetparamSoc('SO_AFFACTFRAISDEFAUT')='')
        or (GetparamSoc('SO_AFFACTFOURDEFAUT')='')
        or (GetparamSoc('SO_AFFACTRESSDEFAUT')='') )
           then begin
                SetParamSOc('SO_AFFACTPARRES',AfFactParREs); //remise ancienne valeur
                PGIInfo('Vos paramètres ne sont pas tous renseignés',titre);
                Lasterror:=1;
                Erreur:=True;
                exit;
           end;
   (* If (GetControlText('AFFACTPARRES1') <> 'SAN')and ((GetParamsoc('So_AFFACTMODEPEC')='DA') or  (GetParamsoc('So_AFFACTMODEPEC')='DF')) then begin
                SetParamSOc('SO_AFFACTPARRES',AfFactParREs); //remise ancienne valeur
                PGIInfo('On ne peut pas lancer l''automatisme sur un choix dates précédentes',titre);
                Lasterror:=1;
                Erreur:=True;
                exit;
                end;*)
     SetControlText('SUIVI','Destruction table en cours');
         // dans tous les cas on détruit les enrgt existant dans la table
     ExecuteSql ('Delete from AFCUMUL where ACU_TYPEAC="FAC"');
     if (GetControlText('AFFACTPARRES1') = 'SAN')then exit ;         //traitement tzerminé si pas gestion
     BoucleEclatFact;
     SetControlText('SUIVI','Création table éclatée');
     end;
end ;

// fct qui boucle sur toutes les pièces de la base
// et qui pour chacune des pièces fait l'éclatement.
procedure TOF_FACT_ECLAT.BoucleEclatFact ;
var tobglo , TobPiece,TobDet : Tob;
  qq:Tquery;
  wi ,ii,nbmois: integer;
  y1,Y2,M1,M2,JJ:word ;
  CleDoc : R_CleDoc ;
  Datedeb,dateFin,date: TDateTime;

begin
  TobGlo :=Tob.Create('Liste piece',Nil,-1);
  DateDeb := Idate1900; DateFin:=Idate2099;
  QQ:=OpenSql('SELECT GP_NATUREPIECEG,GP_DATEPIECE,GP_SOUCHE,GP_NUMERO,GP_INDICEG FROM PIECE WHERE GP_NATUREPIECEG="FAC" OR GP_NATUREPIECEG="FRE" OR GP_NATUREPIECEG="AVC"',True);
  if Not QQ.EOF then TobGlo.loadDetailDb( 'Detail','','',QQ,False);
  Ferme(QQ);
  ii:=0;
  if GetParamSoc('So_AFFACTMODEPEC') = 'DE' then begin
     DateDeb:=GetParamSoc('SO_AFDATEDEBCAB');
     DateFin:=GetParamSoc('SO_AFDATEFINCAB');
     end;
  for wi := 0 to TObGlo.detail.count-1 do
      begin
      tobDet:=TobGlo.detail[wi];
      CleDoc.NaturePiece:=TobDEt.getValue('GP_NATUREPIECEG') ;
      CleDoc.DatePiece:=TobDEt.getValue('GP_DATEPIECE') ;
      CleDoc.Souche:=TobDEt.getValue('GP_SOUCHE') ;
      CleDoc.NumeroPiece:=TobDEt.getValue('GP_NUMERO') ;
      CleDoc.Indice:=TobDEt.getValue('GP_INDICEG');
      TOBPiece:=TOB.Create('PIECE',Nil,-1) ;
         // charge une seule facture.. on peut se permettre de tout prendre
      QQ:=OpenSQL('SELECT * FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False),True) ;
      TOBPiece.SelectDB('',QQ) ;
      Ferme(QQ) ;
      // Lecture Lignes
         // charge une seule facture.. on peut se permettre de tout prendre
      QQ:=OpenSQL('SELECT * FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,False)+' ORDER BY GL_NUMLIGNE',True) ;
      TOBPiece.LoadDetailDB('LIGNE','','',QQ,False,True) ;
      Ferme(QQ) ;
      PieceAjouteSousDetail(TOBPiece);
      if GetParamSoc('So_AFFACTMODEPEC') = 'DE' then begin
         // il faut reagrder si la date de facture est comprise dans les dates cabinet.
         // sinon il faut les changer
         // on passe par un ajout ou soustraction d'un jour pour le cas ou exercice pas sur une année civile
         DecodeDate(GetParamSoc('SO_AFDATEFINCAB'),Y1,M1,JJ);
         DecodeDate(GetParamSoc('SO_AFDATEDEBCAB'),Y2,M2,JJ);
         If Y2=Y1 then Nbmois:=M1-M2
            else NbMois:=M1+(12-M2) ;
         while (TobPiece.getvalue('GP_DATEPIECE') > GetParamSoc('SO_AFDATEFINCAB'))or
             (TobPiece.getvalue('GP_DATEPIECE') < GetParamSoc('SO_AFDATEDEBCAB'))
             do begin
             if TobPiece.getvalue('GP_DATEPIECE') > GetParamSoc('SO_AFDATEFINCAB') then begin
                date :=  DebutDeMois(PlusDAte( GetParamSoc('SO_AFDATEFINCAB'),1,'J'));
                SetParamSoc ('SO_AFDATEDEBCAB',Date);
                Date:=PlusDATe(Date ,NbMois,'M');
                SetParamSoc ('SO_AFDATEFINCAB',FindeMois(Date));
                end
              else begin
                Date:=FinDeMois(PlusDAte( GetParamSoc('SO_AFDATEDEBCAB'),-1,'J'));
                SetParamSoc ('SO_AFDATEFINCAB',Date);
                Date:=PlusDate(Date,-NbMois,'M');
                SetParamSoc ('SO_AFDATEDEBCAB',DebutdeMois( date));
                end;
             end;
         end;
      inc (ii); // affichage pièce en cours tous les 10 enrgt
      if ii >10 then begin
         SetControlText('SUIVI',Format ('Pièce en cours %s %s %s', [TobPiece.getValue('GP_NATUREPIECEG'),
                                      TobPiece.getValue('GP_SOUCHE'),TobPiece.getValue('GP_NUMERO') ]));
         ii:=0;
         end;
      MajFactEclat(TobPiece);
      TobPiece.free;
      end;
  if GetParamSoc('So_AFFACTMODEPEC') = 'DE' then begin
     SetParamSoc('SO_AFDATEDEBCAB',DateDeb);
     SEtParamSoc('SO_AFDATEFINCAB',DateFin);
     end;

  TobGlo.free;

end;

Procedure AFLanceFiche_FAct_Eclat;
begin
if ToutSeulAff  then  exit; // quelqu'un d'autre travaille sur la base
try
AGLLanceFiche ('AFF','FACT_ECLAT','','','');
finally   // Pl le 14/10/02 pour debloquer si plantage intempestif
Bloqueur ('AffToutSeul',False);
end;
end;

Initialization
  registerclasses ( [ TOF_FACT_ECLAT ] ) ; 
end.
