{***********UNITE*************************************************
Auteur  ...... : Nicolas CHARRAIX
Créé le ...... : 26/02/2001
Modifié le ... : 28/02/2001
Description .. : contrôles de cohérences de l'application juridique sur divers 
Suite ........ : informations tes que le capital, le nombre des commissaires 
Suite ........ : aux comptes, les nombres de titre...etc.
Mots clefs ... : COHERENCE, CONTRÔLES
*****************************************************************}
unit UCoherence;
//////////////////////////////////////////////////////////////////
interface
//////////////////////////////////////////////////////////////////
uses
   {$IFNDEF EAGLCLIENT}
   db,         
   {$ELSE}

   {$ENDIF}
   utob, HEnt1, sysutils, HMsgBox, HCtrls;
//////////////////////////////////////////////////////////////////
type
   TCoherence = class(TObject)
//////////////////////////////////////////////////////////////////
   public
      procedure ControlNbTitre(sGuidPerDos_p  : string);

      procedure ControlNbVoix(sGuidPerDos_p : string);
      procedure ControlCapital(sGuidPerDos_p : string);
      procedure ControlAssAct(sGuidPerDos_p, sCodeDos_p : string);
      procedure ControlComCo(sGuidPer_p : string);
      procedure ControlInfoDosAnn(sGuidPer_p : string);
      procedure ControlMessage;

   private
      sMessage_c : string;

      procedure AddMessage(sMessage_p : string);
end;
//////////////////////////////////////////////////////////////////
procedure Verifcoherence(sGuidPerDos_p, sCodedos_p : String);
//////////////////////////////////////////////////////////////////
implementation

//////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 31/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure Verifcoherence(sGuidPerDos_p, sCodedos_p : String);
var
    FCoherence_l : TCoherence;
begin
   FCoherence_l := TCoherence.Create;
   FCoherence_l.ControlNbTitre(sGuidPerDos_p);
   FCoherence_l.ControlNbVoix(sGuidPerDos_p);
   FCoherence_l.ControlCapital(sGuidPerDos_p);
   FCoherence_l.ControlAssAct(sGuidPerDos_p, sCodeDos_p);
   FCoherence_l.ControlComCo(sGuidPerDos_p);
   FCoherence_l.ControlMessage;
   FCoherence_l.free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 30/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
PROCEDURE TCoherence.ControlNbTitre(sGuidPerDos_p : string);
var
   OBAnnuLien_l : TOB;
   Np, Us, Pp, TotTitre : integer;
BEGIN
   OBAnnuLien_l := TOB.Create('ANNULIEN', nil, -1);
   OBAnnuLien_l.LoadDetailFromSQL('select sum(ANL_TTNBUS) as ANL_TTNBUS from ANNULIEN ' +
                                  'where ANL_GUIDPERDOS = "' + sGuidPerDos_p + '" and ANL_TTNBUS <> 0');
   Us := OBAnnuLien_l.Detail[0].GetInteger('ANL_TTNBUS');
   OBAnnuLien_l.ClearDetail;

   OBAnnuLien_l.LoadDetailFromSQL('select sum(ANL_TTNBNP) as ANL_TTNBNP from ANNULIEN ' +
                                  'where ANL_GUIDPERDOS = "' + sGuidPerDos_p + '" and ANL_TTNBNP <> 0');
   Np := OBAnnuLien_l.Detail[0].GetInteger('ANL_TTNBNP');
   OBAnnuLien_l.ClearDetail;

   If Us <> Np then
      AddMessage('Le nombre de titres affectés en nue propriété ('+IntToStr(Np)+') n''est pas égal aux titres affectés en usufruit ('+IntToStr(US)+')');

   OBAnnuLien_l.LoadDetailFromSQL('Select sum(ANL_TTNBPP) as ANL_TTNBPP from ANNULIEN ' +
                                  'where ANL_GUIDPERDOS = "' + sGuidPerDos_p + '" and ANL_TTNBPP <> 0');
   Pp := OBAnnuLien_l.Detail[0].GetInteger('ANL_TTNBPP');
   OBAnnuLien_l.ClearDetail;
   TotTitre := Pp + Np;

   OBAnnuLien_l.LoadDetailFromSQL('Select ANN_CAPNBTITRE, ANN_FORME from ANNUAIRE, JURIDIQUE ' +
                                  'where ANN_GUIDPER = "' + sGuidPerDos_p + '"');

   If (OBAnnuLien_l.Detail[0].GetString('ANN_FORME') <> 'ASS') and
      (TotTitre <> OBAnnuLien_l.Detail[0].GetInteger('ANN_CAPNBTITRE')) then
      AddMessage('Le nombre total de titres affectés aux associés ('+IntToStr(TotTitre)+
             ') n''est pas égal au nombre de titres composant le capital ('+
             OBAnnuLien_l.Detail[0].GetString('ANN_CAPNBTITRE') +')');
   OBAnnuLien_l.Free;
END;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 30/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
PROCEDURE TCoherence.ControlNbVoix(sGuidPerDos_p : string);
var
   OBAnnuLien_l : TOB;
   VAgo, VAge : integer;
BEGIN
   OBAnnuLien_l := TOB.Create('ANNULIEN', nil, -1);
   OBAnnuLien_l.LoadDetailFromSQL('select sum(ANL_VOIXAGO) as ANL_VOIXAGO from ANNULIEN ' +
                                  'where ANL_GUIDPERDOS = "' + sGuidPerDos_p + '" and ANL_VOIXAGO <> 0');

   VAgo := OBAnnuLien_l.Detail[0].GetInteger('ANL_VOIXAGO');
   OBAnnuLien_l.ClearDetail;

   OBAnnuLien_l.LoadDetailFromSQL('Select JUR_NBDROITSVOTE from JURIDIQUE ' +
                                  'where JUR_GUIDPERDOS = "' + sGuidPerDos_p + '" and JUR_TYPEDOS="STE" and JUR_NOORDRE = 1');

   If VAgo <> OBAnnuLien_l.Detail[0].GetInteger('JUR_NBDROITSVOTE') then
      AddMessage('Le nombre total de voix aux AGO affectés aux associés ('+IntToStr(VAgo)+
             ') n''est pas égal au nombre total de droits de vote correspondant au capital ('+
             OBAnnuLien_l.Detail[0].GetString('JUR_NBDROITSVOTE') + ')');
   OBAnnuLien_l.ClearDetail;

   OBAnnuLien_l.LoadDetailFromSQL('select sum(ANL_VOIXAGE) as ANL_VOIXAGE from ANNULIEN ' +
                                  'where ANL_GUIDPERDOS = "' + sGuidPerDos_p + '" and ANL_VOIXAGE <> 0');
   VAge := OBAnnuLien_l.Detail[0].GetInteger('ANL_VOIXAGE');
   OBAnnuLien_l.ClearDetail;

   OBAnnuLien_l.LoadDetailFromSQL('Select JUR_NBDROITSVOTE from JURIDIQUE ' +
                                  'where JUR_GUIDPERDOS = "' + sGuidPerDos_p + '" and JUR_TYPEDOS="STE" and JUR_NOORDRE = 1');

   If VAge <> OBAnnuLien_l.Detail[0].GetInteger('JUR_NBDROITSVOTE') then
      AddMessage('Le nombre total de voix aux AGE affectés aux associés ('+IntToStr(VAge)+
             ') n''est pas égal au nombre total de droits de vote correspondant au capital ('+
             IntToStr(OBAnnuLien_l.Detail[0].GetInteger('JUR_NBDROITSVOTE')) + ')');
   OBAnnuLien_l.Free;
END;
 


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 30/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
PROCEDURE TCoherence.ControlCapital(sGuidPerDos_p : string);
var
   OBJuridique_l, OBJurForme_l : TOB;
   capdev :string;

BEGIN
   OBJurForme_l := TOB.Create('JUFORMEJUR', nil, -1);
   OBJurForme_l.LoadDetailFromSQL(
               'Select JUR_CODEDOS, JUR_FORME, JUR_CAPDEV, JUR_APPELPUB, JUR_CAPITAL, JUR_CAPDEV, ' +
               'JFJ_FORME, JFJ_CONTROLE , JFJ_CAPDEVISE, JFJ_CAPMINAPE, JFJ_CAPDEVISE, JFJ_CAPMIN ' +
               'FROM JURIDIQUE, JUFORMEJUR ' +
               'Where JUR_GUIDPERDOS = "' + sGuidPerDos_p + '" and JUR_FORME = JFJ_FORME and JFJ_CONTROLE = "X"');

   //LIEN SUR JUR_FORME(dossier_en_cours) = JFJ_FORME
   //A EXECUTER SEULEMENT SI JFJ_CONTROLE ="X"
   if OBJurForme_l.Detail.Count = 0 then
   begin
      OBJurForme_l.Free;
      exit;
   end;

   //CAPITAL
   If OBJurForme_l.Detail[0].GetString('JUR_CAPDEV') <> OBJurForme_l.Detail[0].GetString('JFJ_CAPDEVISE')  then   //// (= pas la même devise entre cap. société et cap. mini pour forme)
   begin
      OBJuridique_l := TOB.Create('JURIDIQUE', nil, -1);
      OBJuridique_l.LoadDetailDBFromSQL('JURIDIQUE',
                                        'Select JUR_CAPDEV From JURIDIQUE ' +
                                        'Where JUR_CODEDOS = "' + OBJurForme_l.Detail[0].GetString('JUR_CODEDOS') + '"');
      capdev := '';
      If (OBJurForme_l.Detail[0].GetString('JUR_CAPDEV') <> '' ) and
         (OBJurForme_l.Detail[0].GetString('JFJ_CAPDEVISE') <> '') then
      begin
         If (OBJurForme_l.Detail[0].GetString('JUR_CAPDEV') ='FRF') AND
            (OBJurForme_l.Detail[0].GetString('JFJ_CAPDEVISE') = 'EUR')  then
         begin
            capdev := 'EUR';
         end
         else If(OBJurForme_l.Detail[0].GetString('JUR_CAPDEV') ='EUR') AND
                (OBJurForme_l.Detail[0].GetString('JFJ_CAPDEVISE') = 'FRF') then
         begin
            capdev := 'FRF';
         end;

         if capdev <>'' then
         begin
            OBJuridique_l.Detail[0].PutValue('JUR_CAPDEV', capdev);
            OBJuridique_l.Detail[0].UpdateDB;
         end;
      end;
     OBJuridique_l.Free;
   end
   else  //(devise capital = devise pour la forme)
   begin
      If OBJurForme_l.Detail[0].GetString('JUR_APPELPUB') = 'X' then //  (= appel public à l'épargne)
      begin
         If (OBJurForme_l.Detail[0].GetInteger('JFJ_CAPMINAPE') <> 0)  and
            (OBJurForme_l.Detail[0].GetInteger('JUR_CAPITAL') < OBJurForme_l.Detail[0].GetInteger('JFJ_CAPMINAPE')) then
            AddMessage('Le capital de la société ('+OBJurForme_l.Detail[0].GetString('JUR_CAPITAL') + ' ' + OBJurForme_l.Detail[0].GetString('JUR_CAPDEV') + ') est inférieur au capital'+
                    ' minimum pour cette forme de société ('+OBJurForme_l.Detail[0].GetString('JFJ_CAPMINAPE') + ' ' + OBJurForme_l.Detail[0].GetString('JFJ_CAPDEVISE') +
                    ') lorsqu''elle fait appel public à l''épargne');
      end
      else //  (= pas d'appel public - cas le plus fréquent)
         If (OBJurForme_l.Detail[0].GetInteger('JFJ_CAPMIN') <> 0) and
            (OBJurForme_l.Detail[0].GetInteger('JUR_CAPITAL') < OBJurForme_l.Detail[0].GetInteger('JFJ_CAPMIN')) then
            AddMessage('Le capital de la société ('+OBJurForme_l.Detail[0].GetString('JUR_CAPITAL') + ' ' + OBJurForme_l.Detail[0].GetString('JUR_CAPDEV') + ') est inférieur au capital'+
                    ' minimum pour cette forme de société ('+OBJurForme_l.Detail[0].GetString('JFJ_CAPMIN') + ' ' + OBJurForme_l.Detail[0].GetString('JFJ_CAPDEVISE') +
                    ')');
   end;
   OBJurForme_l.Free;
END;






{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 30/03/2006
Modifié le ... :   /  /    
Description .. : 2.2 NOMBRE D'ASSOCIES/ACTIONNAIRES
Mots clefs ... : 
*****************************************************************}
PROCEDURE TCoherence.ControlAssAct(sGuidPerDos_p, sCodeDos_p : string);
var
   totAss : integer;
   OBJurForme_l, OBAnnulien_l : TOB;
   sWhere_l : string;
BEGIN
//   if sCodeDos_p <> '&#@' then
//      sWhere_l := '  and JUR_CODEDOS = "' + sCodeDos_p + '" ';

   OBJurForme_l := TOB.Create('JUFORMEJUR', nil, -1);
   OBJurForme_l.LoadDetailFromSQL('Select JFJ_NBASSMIN, JFJ_NBASSMAX ' +
                                  'from JURIDIQUE, JUFORMEJUR ' +
                                  'where JUR_GUIDPERDOS = "' + sGuidPerDos_p + '" ' +
                                  sWhere_l +
//                                  'where JUR_CODEDOS = "' + sCodeDos_p + '" ' +
                                  '  and JUR_TYPEDOS = "STE" and JUR_NOORDRE = 1 ' +
                                  '  and JFJ_FORME = JUR_FORME and JFJ_CONTROLE = "X"');

   if OBJurForme_l.Detail.Count = 0 then
   begin
      OBJurForme_l.Free;
      exit;
   end;

//   if sCodeDos_p <> '&#@' then
//      sWhere_l := '  and JUR_CODEDOS = "' + sCodeDos_p + '" and JUR_CODEDOS = ANL_CODEDOS ';

   OBAnnulien_l := TOB.Create('ANNULIEN', nil, -1);
   OBAnnulien_l.LoadDetailFromSQL('select count(ANL_GUIDPER) as ANL_GUIDPER ' +
                                  'from JURIDIQUE, ANNULIEN, JUFONCTION ' +
                                  'where JUR_GUIDPERDOS = "' + sGuidPerDos_p + '" ' +
                                  '  and JUR_GUIDPERDOS = ANL_GUIDPERDOS ' +
                                  sWhere_l +
//                                  'where JUR_CODEDOS = "' + sCodeDos_p + '" ' +
//                                  '  and JUR_CODEDOS = ANL_CODEDOS ' +
                                  '  and JUR_TYPEDOS = "STE" and JUR_NOORDRE = 1 ' +
                                  '  and JFT_FORME = JUR_FORME ' +
                                  '  and JFT_TITRE = "X" and ANL_FONCTION = JFT_FONCTION');

   TotAss := OBAnnulien_l.Detail[0].GetInteger('ANL_GUIDPER');
   OBAnnulien_l.Free;

   If OBJurForme_l.Detail[0].GetInteger('JFJ_NBASSMIN') <> 0 then
      If TotAss < OBJurForme_l.Detail[0].GetInteger('JFJ_NBASSMIN') then
         AddMessage('Le nombre d''associés saisis ('+InttoStr(TotAss) +
                 ') est inférieur au minimum prévu pour cette forme de société (' +
                 IntToStr(OBJurForme_l.Detail[0].GetInteger('JFJ_NBASSMIN'))+')');

   If OBJurForme_l.Detail[0].GetInteger('JFJ_NBASSMAX') <> 0 then
      If TotAss > OBJurForme_l.Detail[0].GetInteger('JFJ_NBASSMAX') then
         AddMessage('Le nombre d''associés saisis (' + InttoStr(TotAss)+
                  ') est supérieur au maximum prévu pour cette forme de société ('+
                  IntToStr(OBJurForme_l.Detail[0].GetInteger('JFJ_NBASSMAX'))+')');
   OBJurForme_l.Free;
END;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 30/03/2006
Modifié le ... :   /  /    
Description .. : 2.3    NOMBRE DE COMMISSAIRES AUX COMPTES
Suite ........ : 2.3.1  Sur les commissaires existants dans le dossier
Mots clefs ... : 
*****************************************************************}
PROCEDURE TCoherence.ControlComCo(sGuidPer_p : string);
Var
   OBAnnuLien_l : TOB;
TotCCT, TotCCS : integer;
BEGIN
   OBAnnuLien_l := TOB.Create('ANNULIEN', nil, -1);
   OBAnnuLien_l.LoadDetailFromSQL('Select ANL_GUIDPERDOS from ANNULIEN, JUFORMEJUR ' +
                                  'where ANL_GUIDPERDOS = "' + sGuidPer_p + '" ' +
                                  '  and ANL_FORME=JFJ_FORME And JFJ_CONTROLE = "X"');
   if OBAnnuLien_l.Detail.Count = 0 then
   begin
      OBAnnuLien_l.Free;
      exit;
   end;

   OBAnnuLien_l.ClearDetail;
   OBAnnuLien_l.LoadDetailFromSQL('Select count(ANL_GUIDPERDOS) as ANL_GUIDPERDOS from ANNULIEN ' +
                                  'where ANL_GUIDPERDOS = "' + sGuidPer_p + '" and ANL_FONCTION = "CCT"');
   TotCCT := OBAnnuLien_l.Detail[0].GetInteger('ANL_GUIDPERDOS');
   OBAnnuLien_l.ClearDetail;

   OBAnnuLien_l.LoadDetailFromSQL('Select count(ANL_GUIDPERDOS) as ANL_GUIDPERDOS from ANNULIEN ' +
                                  'where ANL_GUIDPERDOS = "' + sGuidPer_p + '" and ANL_FONCTION = "CCS"');
   TotCCS := OBAnnuLien_l.Detail[0].GetInteger('ANL_GUIDPERDOS');

   if TotCCT <> TotCCS  then
     AddMessage('Le nombre de commissaire(s) aux comptes titulaire(s) ('+InttoStr(TotCCT)+') est différent du nombre de commissaire(s) suppléant(s) ('+InttoStr(TotCCS)+')');

   //2.3.2 Sur le nombre minimum légal de commissaires
   OBAnnuLien_l.ClearDetail;
   OBAnnuLien_l.LoadDetailFromSQL('Select JFJ_NBCCT from JUFORMEJUR, JURIDIQUE ' +
                                  'WHERE JUR_GUIDPERDOS = "' + sGuidPer_p + '" and JUR_FORME = JFJ_FORME');
   If (OBAnnuLien_l.Detail.Count > 0)
    and (OBAnnuLien_l.Detail[0].GetInteger('JFJ_NBCCT') <> 0) then
      If TotCCT < OBAnnuLien_l.Detail[0].GetInteger('JFJ_NBCCT') then
         AddMessage('Le nombre de commissaire(s) aux comptes titulaire(s) ('+InttoStr(TotCCT)+') est inférieur au nombre prescrit pour cette forme de société ('+IntToStr(OBAnnuLien_l.Detail[0].GetInteger('JFJ_NBCCT'))+')');
   OBAnnuLien_l.Free;
END;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 30/03/2006
Modifié le ... :   /  /    
Description .. : 3. CONTROLE SUR INFORMATIONS
Suite ........ : STATUTAIRES
Suite ........ : LIEN SUR JUR_FORME(dossier_en_cours) = JFJ_FORME
Suite ........ :
Suite ........ : 4. CONTROLE SUR LIMITES D'AGE DES
Suite ........ : MANDATAIRES   
Suite ........ : 
Suite ........ : conctrole de cohérence entre les infos de l'annuaire et celle 
Suite ........ : du dossier juridique.
Suite ........ : appelée nul part et sans aucun effet pour l'instant, à 
Suite ........ : réfléchir......
Mots clefs ... : 
*****************************************************************}
procedure TCoherence.ControlInfoDosAnn(sGuidPer_p: string);
var
   OBAnnuLien_l: TOB;
ListeChamp, TempLstA,TempLstJ,select,champ,champA, ChampJ, LibChamp : string ;

begin
   ListeChamp := 'CAPITAL;CAPDEV;MOISCLOTURE;FORME;CAPNBTITRE;CAPVN';
   TempLstA := ListeChamp;
   select :='';
   While TempLstA<>'' do
   begin
      champ := ReadTokenSt(TempLstA);
      ChampA := 'ANN_'+Champ;
      ChampJ := 'JUR_'+Champ;
      Select := select+champA+' AS A'+Champ+',';
   end;
   
   OBAnnuLien_l := TOB.Create('ANNULIEN', nil, -1);
   OBAnnuLien_l.LoadDetailFromSQL(Copy(Select,1,Length(Select)-1) + ' FROM ANNUAIRE, JURIDIQUE ' +
                                 'WHERE ANN_GUIDPER = "' + sGuidPer_p + '" AND JUR_GUIDPERDOS = ANN_GUIDPER');
   If OBAnnuLien_l.detail.Count > 0 then
   begin
      TempLstJ := ListeChamp;
      While TempLstJ<>'' do
      begin
         champ := ReadTokenSt(TempLstJ);
         If OBAnnuLien_l.Detail[0].GetValue('J'+Champ) <> OBAnnuLien_l.Detail[0].GetValue('A'+Champ) then
         begin
            if champ='CAPITAL' then LibChamp := 'Le capital'
            else if champ='CAPDEV' then LibChamp := 'La devise du capital'
            else if champ='MOISCLOTURE' then LibChamp := 'Le Mois de cloture'
            else if champ='FORME' then  Libchamp := 'La forme'
            else if champ='CAPNBTITRE' then  Libchamp := 'Le nombre total des droits de vote'
            else if champ='CAPVN' then Libchamp := 'La valeur nominale des titres';
            AddMessage(libchamp+' du dossier en cours est différent du +'+Libchamp+' de la personne rattaché');
         end;
      end;
   end;
   OBAnnuLien_l.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 31/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TCoherence.AddMessage(sMessage_p : string);
begin
   if sMessage_c <> '' then
      sMessage_c := sMessage_c + #13#10;
   sMessage_c := sMessage_c + '- ' + sMessage_p;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 31/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TCoherence.ControlMessage;
begin
   if sMessage_c <> '' then
      PGIInfo('ATTENTION : #13#10' + sMessage_c, TitreHalley + '  - Contrôle de cohérence');
end;
//////////////////////////////////////////////////////////////////
end.
