unit UtilRCC;

interface

uses  Sysutils, HCtrls, HEnt1, Controls,
{$IFDEF EAGLCLIENT}
{$ELSE}
      {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}                                
      HMsgBox,Classes, UTob, YRessource;

Procedure RTIntegrationRessourcePaie ;
// duplication tablettes paie dans ressources : mng 21-05-02
Procedure RTCopieTablettes ( CodeOrigine,CodeDestination : String);

implementation

{ Utilitaire intégration salariés dans ables RESSOURCES et Salariés }
Procedure RTIntegrationRessourcePaie ;
var requete,MessRCC : String;
    NbRessourceCrees : Integer;
    TobRessource,TR : TOB;
    Q,QQ : TQuery;
    i : Integer;
//    Etablissement principal;Matricule Salarié;Code Ressource;Nom du salarié;Prénom;Date d'entrée;Date de sortie;Libellé Emploi
begin
  if PGIAsk('Voulez-vous intégrer les Ressources Paie ?', 'Intégration') <> mrYes then exit;

  NbRessourceCrees:=0;
  TobRessource:=Tob.create('les ressources',Nil,-1);

  try
    // ajout tablettes : mng 21-05-02
    requete:='Select PSA_ETABLISSEMENT,PSA_SALARIE, PSA_LIBELLE, PSA_PRENOM, PSA_LIBELLEEMPLOI'+
             ', PSA_TRAVAILN1, PSA_TRAVAILN2, PSA_TRAVAILN4, PSA_LIBREPCMB1, PSA_LIBREPCMB2, PSA_LIBREPCMB3, PSA_LIBREPCMB4 FROM SALARIES';
    Q:=OpenSQL(requete,True);

    while not Q.Eof do
    begin
      TR:=Tob.create('RESSOURCE',TobRessource,-1);
      TR.InitValeurs;
      TR.PutValue('ARS_TYPERESSOURCE', 'SAL' );

      TR.PutValue('ARS_RESSOURCE', IntToStr(StrToInt(Q.FindField('PSA_SALARIE').AsString)));
      TR.PutValue('ARS_ETABLISSEMENT', Q.FindField('PSA_ETABLISSEMENT').AsString );
      TR.PutValue('ARS_SALARIE', Q.FindField('PSA_SALARIE').AsString );

      //Recup du code utilisateur dans la table utilisat
      QQ:=OpenSQL('Select US_UTILISATEUR From UTILISAT where US_AUXILIAIRE="'+Q.FindField('PSA_SALARIE').AsString+'"',True);
      if not QQ.Eof then
         TR.PutValue('ARS_UTILASSOCIE', QQ.FindField('US_UTILISATEUR').AsString );
      Ferme(QQ);

      TR.PutValue('ARS_LIBELLE', Q.FindField('PSA_LIBELLE').AsString );
      TR.PutValue('ARS_LIBELLE2', Q.FindField('PSA_PRENOM').AsString );
      TR.PutValue('ARS_FONCTION1', Q.FindField('PSA_LIBELLEEMPLOI').AsString );

  // ajout tablettes : mng 21-05-02
      TR.PutValue('ARS_LIBRERES1', Q.FindField('PSA_TRAVAILN1').AsString );
      TR.PutValue('ARS_LIBRERES2', Q.FindField('PSA_LIBREPCMB1').AsString );
      TR.PutValue('ARS_LIBRERES3', Q.FindField('PSA_LIBREPCMB3').AsString );
      TR.PutValue('ARS_LIBRERES4', Q.FindField('PSA_LIBREPCMB2').AsString );
      TR.PutValue('ARS_LIBRERES5', Q.FindField('PSA_LIBREPCMB4').AsString );
      TR.PutValue('ARS_LIBRERES6', Q.FindField('PSA_TRAVAILN2').AsString );
      TR.PutValue('ARS_LIBRERES7', Q.FindField('PSA_TRAVAILN4').AsString );
  // mng 04/12/2002
      TR.PutValue('ARS_DEPARTEMENT', Q.FindField('PSA_LIBREPCMB4').AsString );

      Inc (NbRessourceCrees);
      Q.Next;
    end;

    Ferme(Q);

    // C.B 03/04/2007                    
    // hamonisation des plannings
    if TobRessource.InsertorUpdateDB(true) then
      for i := 0 to TobRessource.detail.count - 1 do
        CreateYRS(TobRessource.detail[i].GetString('ARS_SALARIE'), TobRessource.detail[i].GetString('ARS_RESSOURCE'), TobRessource.detail[i].GetString('ARS_UTILASSOCIE'));
                                        
  finally
    TobRessource.Free;
  end;

  // duplication tablettes paie dans ressources : mng 21-05-02
  RTCopieTablettes ('PAG','LR1');
  RTCopieTablettes ('PL1','LR2');
  RTCopieTablettes ('PL3','LR3');
  RTCopieTablettes ('PL2','LR4');
  RTCopieTablettes ('PL4','LR5');
  RTCopieTablettes ('PST','LR6');
  RTCopieTablettes ('PSV','LR7');
  // mng 04-12-2002
  RTCopieTablettes ('PL4','ADP');

  MessRCC:='Nombre de Ressource créés : '+IntToStr(NbRessourceCrees);
  PGIBox(MessRCC,'Informations sur l''intégration');

end;

// duplication tablettes paie dans ressources : mng 21-05-02
Procedure RTCopieTablettes ( CodeOrigine,CodeDestination : String);
var TobCode,TobChoixExt,TobExt : TOB;
    Q : TQuery;
    i,j : integer;
    FieldNameFrom,FieldNameTo : string;
begin
Q:=OpenSQL('Select * from CHOIXCOD where CC_TYPE="'+CodeOrigine+'"',True);
//TobCode:=Nil;
TobCode:=TOB.create ('table choixcod',NIL,-1);
TobCode.LoadDetailDB('CHOIXCOD','','',Q, false);
Ferme (Q);

TobExt:=TOB.create ('table choixext',NIL,-1);


for i := 0 to TobCode.Detail.Count-1 do
    begin
       TobChoixExt:=Tob.create('CHOIXEXT',TobExt,-1);
       TobChoixExt.InitValeurs;
       for j :=1 to TobCode.detail[i].NbChamps do
          begin
          FieldNameFrom := TobCode.detail[i].GetNomChamp(j);
          FieldNameTo  := 'YX_' + copy(FieldNameFrom,4,length(FieldNameFrom)) ;
          TobChoixExt.PutValue(FieldNameTo, TobCode.detail[i].GetValue(FieldNameFrom));
          TobChoixExt.putvalue('YX_TYPE',CodeDestination);
          end;
    end;
TobExt.InsertOrUpdateDB(False);

TobExt.Free ;
TobCode.Free ;
end;

end.


