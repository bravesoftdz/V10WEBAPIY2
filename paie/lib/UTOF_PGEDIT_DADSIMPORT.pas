{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/07/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDIT_DADSIMPORT ()
Mots clefs ... : TOF;PGEDIT_DADSIMPORT
*****************************************************************}
Unit UTOF_PGEDIT_DADSIMPORT;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
     QRS1,
{$else}
     eMul,
     eQRS1,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     PgDADSOutils,
     Hstatus,
     ed_tools,
     PGDADSControles,
     PGDADSCommun,
     UTobDebug,
     paramsoc;

Type
  TOF_PGEDIT_DADSIMPORT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    TobLexique, TobPeriodes : Tob;
    PerDeb, PerFin : string;
    procedure RecupFichier ();
    function AffecteDebut(T : TOB) : integer;
    function AffecteFin(T : TOB) : integer;
    procedure ChargerLexiqueDads (Annee, NatureDADSU:String);
    procedure CreationSegments (ExerciceDads:string; var T:tob; TobEntete:tob);
    procedure EnregistreSegment(TL : tob; ExerciceDads, Segment, ValeurSeg : string;var T : tob);
  end ;

Implementation

procedure TOF_PGEDIT_DADSIMPORT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGEDIT_DADSIMPORT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGEDIT_DADSIMPORT.OnUpdate;
begin
Inherited;
FreeAndNil (TDADSUD);
ChargeTOBDADS;
try
   begintrans;
   RecupFichier;
{ggg
   InitMoveProgressForm (NIL,'Edition en cours', 'Veuillez patienter SVP ...',
                         TDADSUD.FillesCount (4),FALSE,TRUE);
   InitMove(TDADSUD.FillesCount (4), '');
   FiniMove;
   FiniMoveProgressForm;
}   
   CommitTrans;
except
   Rollback;
   end;

TobDebug (TDADSUD);
FreeTobDADS;
FreeAndNil (TLexique);

TobDebug (TobPeriodes);
TFQRS1(Ecran).LaTob:= TobPeriodes;

end ;

procedure TOF_PGEDIT_DADSIMPORT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGEDIT_DADSIMPORT.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_PGEDIT_DADSIMPORT.OnClose ;
begin
Inherited;
FreeAndNil (TDADSUD);
end ;

procedure TOF_PGEDIT_DADSIMPORT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGEDIT_DADSIMPORT.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/07/2006
Modifié le ... :   /  /
Description .. : Procédure de remplissage des tables de la DADS
Mots clefs ... : PAYE;PGDADSU
*****************************************************************}
procedure TOF_PGEDIT_DADSIMPORT.RecupFichier ();
var
DateDeb, DateDebMMJJ, DateFin, DateFinMMJJ, ExerciceDads, FileN : String;
Fraction, Libelle, NatureDADSU, NIC, Nom, Periodicite, Prenom, S : String;
Salarie, Segment, Structure, ValeurSeg : String;
FLect : TextFile;
T, TErreurDetail, TL, TobEmetteur, TobEmetteurD, TobEntete, TobEntreprise : TOB;
TobEtabliss, TobFille, TobHonoraire, TobSalaries, TobTotaux, TobTotauxD : TOB;
TobUneEntreprise, TobUneEntrepriseD, TobUnEtab, TobUnEtabD : TOB;
TobUnHonoraire, TobUnHonoraireD, TobUnSalarie, TobUnSalarieD : TOB;
i, IncBBS, MaxBBS, NbBBS, OrdreBBS, IncBPS, MaxBPS, NbBPS, OrdreBPS, IncExo, MaxExo, NbExo, OrdreExo, IncPrev, MaxPrev, NbPrev, OrdrePrev, IncRC, MaxRC, NbRC, OrdreRC, IncTSI, MaxTSI, NbTSI, OrdreTSI : integer;
LL, Ordre, Virgule : integer;
CreationTob : Boolean;
Q : TQuery;
begin
//Récupération du fichier
FileN:= GetControlText ('FICHIER');

if (FileN='') then
   begin
   PGIBox ('Aucun fichier sélectionné', Ecran.Caption);
   exit;
   end;

AssignFile (FLect, FileN);
Reset (FLect);

TDADSUD:= Tob.Create ('DADSDETAIL', Nil, -1);
TobEmetteur:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
TobEntreprise:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
TobSalaries:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
TobHonoraire:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
TobEtabliss:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
TobTotaux:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
TLexique:= Tob.Create ('DADSLEXIQUE', Nil, -1);
TobLexique:= Tob.Create ('LesSegments', Nil, -1);

TobPeriodes:= Tob.Create('LesPeriodes',Nil,-1);

ExerciceDads:= '';
CreationTob:= False;

Readln (FLect,S);
LL:= Length (S);
Virgule:= Pos(',', S);
Segment:= Copy (S,1,Virgule-1);
Structure:= Copy (S,1,3);
ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));

while (not(eof(FLect))) do
      begin
      if (Segment = 'S10.G01.00.001.001') then
         begin
         Ordre:= 0;
         Libelle:= '';
         TypeD:= '';

         while (Structure = 'S10') do
               begin
               TobEmetteurD:= Tob.Create ('DADSDETAIL', TobEmetteur, -1);
               if (Segment='S10.G01.00.002') then
                  begin
                  Libelle:= ValeurSeg;
                  TobEmetteur.PutValueAllFille ('PDS_SALARIE', Libelle);
                  end
               else
               if (Libelle <> '') then
                  TobEmetteurD.PutValue ('PDS_SALARIE', Libelle);

               if (Segment='S10.G01.01.001.001') then
                  Ordre:= Ordre+1;
               TobEmetteurD.PutValue ('PDS_ORDRE', Ordre);
               TobEmetteurD.PutValue ('PDS_SEGMENT', Segment);
               TobEmetteurD.PutValue ('PDS_DONNEE', ValeurSeg);
               Readln (FLect,S);
               LL:= Length (S);
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
               Structure:= Copy (S,1,3);
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
               end;
         end;

      if (Segment = 'S20.G01.00.001') then
         begin
         Ordre:= 0;
         Libelle:= '';
         TobUneEntreprise:= Tob.Create ('DADSDETAIL', TobEntreprise, -1);

         while (Structure = 'S20') do
               begin
               TobUneEntrepriseD:= Tob.Create ('DADSDETAIL', TobUneEntreprise, -1);
               if (Segment = 'S20.G01.00.002') then
                  begin
                  Libelle:= ValeurSeg;
                  TobUneEntreprise.PutValueAllFille ('PDS_SALARIE', Libelle);
                  SetControlText ('DOSSIER', Libelle);
                  end
               else
               if (Libelle <> '') then
                  TobUneEntrepriseD.PutValue ('PDS_SALARIE', Libelle);

               If (Segment='S20.G01.00.003.001') then
                  begin
                  DateDeb:= Copy (ValeurSeg, 1, 2)+'/'+Copy (ValeurSeg, 3, 2)+
                            '/'+Copy (ValeurSeg, 5, 4);
                  DateDebMMJJ:= Copy (ValeurSeg, 3, 2)+Copy (ValeurSeg, 1, 2);
                  end;

               if (Segment='S20.G01.00.003.002') then
                  begin
                  DateFin:= Copy (ValeurSeg, 1, 2)+'/'+Copy (ValeurSeg, 3, 2)+
                            '/'+Copy (ValeurSeg, 5, 4);
                  DateFinMMJJ:= Copy (ValeurSeg, 3, 2)+Copy (ValeurSeg, 1, 2);
                  ExerciceDads:= Copy (ValeurSeg, 5, 4);
                  PGAnnee:= ExerciceDads;
                  end;

               if (Segment='S20.G01.00.004.001') then
               //Recherche du type de DADS
                  begin
                  NatureDADSU:= ValeurSeg;
                  if (NatureDADSU='04') then
                     ExerciceDads:= IntToStr(StrToInt(ExerciceDads)-1);
                  end;

               if (Segment='S20.G01.00.004.002') then
               //Recherche du type de DADS
                  begin
                  if (ValeurSeg='52') then
                     TypeD:= '2'+Copy (TypeD, 2, 2);

                  if (ValeurSeg='55') then
                     SetControlChecked ('CHNEANT', True);
                  end;

               If (Segment='S20.G01.00.005') then
                  Fraction:= Copy (ValeurSeg, 1, 1);

               If (Segment='S20.G01.00.018') then
               // Recherche de la périodicité de la DADS
                  begin
                  if (NatureDADSU<>'04') then
                     TypeD:= Copy (TypeD, 1, 1)+'01'
                  else
                  if ((Copy (ValeurSeg, 1, 1)='A')) then
                     TypeD:= Copy (TypeD, 1, 1)+'02'
                  else
                  if ((Copy (ValeurSeg, 1, 1)='S')) then
                     TypeD:= Copy (TypeD, 1, 1)+'03'
                  else
                  if ((Copy (ValeurSeg, 1, 1)='T')) then
                     TypeD:= Copy (TypeD, 1, 1)+'04'
                  else
                  if ((Copy (ValeurSeg, 1, 1)='M')) then
                     TypeD:= Copy (TypeD, 1, 1)+'05';
                  Periodicite:= ValeurSeg;
                  end;

               TobUneEntrepriseD.PutValue ('PDS_ORDRE', Ordre);
               TobUneEntrepriseD.PutValue ('PDS_SEGMENT', Segment);
               TobUneEntrepriseD.PutValue ('PDS_DONNEE', ValeurSeg);
               Readln (FLect,S);
               LL:= Length (S);
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
               Structure:= Copy (S,1,3);
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
               end;
         if (Structure>'S20') then
            begin
            TobEmetteur.PutValueAllFille ('PDS_TYPE', TypeD);
            TobEmetteur.PutValueAllFille ('PDS_EXERCICEDADS', ExerciceDads);
            TobEmetteur.PutValueAllFille ('PDS_DATEDEBUT', StrToDate (DateDeb));
            TobEmetteur.PutValueAllFille ('PDS_DATEFIN', StrToDate (DateFin));
            TobUneEntreprise.PutValueAllFille ('PDS_TYPE', TypeD);
            TobUneEntreprise.PutValueAllFille ('PDS_EXERCICEDADS',
                                               ExerciceDads);
            TobUneEntreprise.PutValueAllFille ('PDS_DATEDEBUT',
                                               StrToDate (DateDeb));
            TobUneEntreprise.PutValueAllFille ('PDS_DATEFIN',
                                               StrToDate (DateFin));
            SetControlText ('DATEDEB', DateDeb);
            SetControlText ('DATEFIN', DateFin);
            end;
         end;

{Construction tob qui contient les définitions des champs correspondant aux
segments à partir table DADSLEXIQUE}
      if ((ExerciceDads<>'') and (NatureDADSU<>'') and
         (TLexique.Detail.Count=0)) then
         ChargerLexiqueDads(ExerciceDads, NatureDADSU);

      if (Segment='S30.G01.00.001') then
      //Si nouveau salarié
         begin
         Ordre:= 0;
         Libelle:= '';
         Nom:= '';
         Prenom:= '';
         TobUnSalarie:= Tob.Create ('DADSDETAIL', TobSalaries, -1);
{££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££}
         TobEntete:= Tob.Create('MonEntete',Nil,-1);
{******************************************************************************}

         while (Structure = 'S30') do
               begin
               TobUnSalarieD:= Tob.Create ('DADSDETAIL', TobUnSalarie, -1);
               TobUnSalarieD.PutValue ('PDS_TYPE', TypeD);
               TobUnSalarieD.PutValue ('PDS_ORDRE', Ordre);
               TobUnSalarieD.PutValue ('PDS_SEGMENT', Segment);
               TobUnSalarieD.PutValue ('PDS_DONNEE', ValeurSeg);
               TobUnSalarieD.PutValue ('PDS_DATEDEBUT', StrToDate (DateDeb));
               TobUnSalarieD.PutValue ('PDS_DATEFIN', StrToDate (DateFin));
               TobUnSalarieD.PutValue ('PDS_EXERCICEDADS', ExerciceDads);
               if (Segment = 'S30.G01.00.002') then
                  Nom:= ValeurSeg
               else
               if (Segment = 'S30.G01.00.003') then
                  Prenom:= ValeurSeg
               else
               if (Segment = 'S30.G01.00.004') then
                  Nom:= ValeurSeg;
{££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££}
               TobEntete.AddchampSupValeur(Segment, ValeurSeg, False);
{******************************************************************************}
               Readln (FLect,S);
               LL:= Length (S);
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
               Structure:= Copy (S,1,3);
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
               end;
         end;

      if (Segment='S41.G01.00.001') then
      //Si nouvelle période
         begin
         PerDeb:= '';
         PerFin:= '';

         while ((Structure = 'S41') or (Structure = 'S42') or
               (Structure = 'S43') or (Structure = 'S44') or
               (Structure = 'S45') or (Structure = 'S46') or
               (Structure = 'S51') or (Structure = 'S53') or
               (Structure = 'S66')) do
               begin
               if ((Segment = 'S41.G01.00.001') and (Ordre <>0)) then
                  begin
                  TobUnSalarie.ParcoursTraitement (['PDS_ORDRE'], [Ordre],
                                                   False, AffecteDebut);
                  TobUnSalarie.ParcoursTraitement (['PDS_ORDRE'], [Ordre],
                                                   False, AffecteFin);
                  end;

               TobUnSalarieD:= Tob.Create ('DADSDETAIL', TobUnSalarie, -1);

               if (Segment = 'S41.G01.00.001') then
                  begin
                  Ordre:= Ordre+1;
                  PerDeb:= Copy (ValeurSeg, 3, 2)+Copy (ValeurSeg, 1, 2);
                  if (((DateDebMMJJ<=PerDeb) and (PerDeb<=DateFinMMJJ)) or
                     (PerDeb<=DateDebMMJJ)) then
                     PerDeb:= Copy (ValeurSeg, 1, 2)+'/'+Copy (ValeurSeg, 3, 2)+
                                    '/'+ExerciceDADS
                  else
                     PerDeb:= Copy (ValeurSeg, 1, 2)+'/'+Copy (ValeurSeg, 3, 2)+
                                    '/'+IntToStr(StrToInt(ExerciceDads)-1);

////////////////////////////////////////////////////////////////////////////////
                  T:= Tob.create('Filles',TobPeriodes,-1);
                  T.AddchampSupValeur('DOSSIER', GetControlText('DOSSIER'), False);
                  T.AddchampSupValeur('DATEDEB', GetControlText('DATEDEB'), False);
                  T.AddchampSupValeur('DATEFIN', GetControlText('DATEFIN'), False);

                  if (NatureDADSU='04') then
                     T.AddchampSupValeur('BTP', 'BTP', False)
                  else
                     T.AddchampSupValeur('BTP', '', False);

                  CreationSegments(ExerciceDads, T, TobEntete);
                  CreationTob:= True;
////////////////////////////////////////////////////////////////////////////////
                  end;
{££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££}
               if (CreationTob=True) then
                  begin
                  Segment:= Copy (S,1,14);
                  LL:= Length (S);
                  ValeurSeg:= Copy (S, 17, LL-17)+ExerciceDads;
                  T.PutValue (Segment,StrToDate(Copy (ValeurSeg,1,2)+'/'+
                              Copy (ValeurSeg,3,2)+'/'+Copy (ValeurSeg,5,4)));
                  Readln (FLect,S);
                  Structure:= Copy (S, 1, 3);
                  Virgule:= Pos(',', S);
                  Segment:= Copy (S,1,Virgule-1);
                  OrdreRC:= 101;
                  NbRC:= 0;
                  IncRC:= 2;
                  MaxRC:= 109;
                  OrdreBBS:= 201;
                  NbBBS:= 0;
                  IncBBS:= 3;
                  MaxBBS:= 204;
                  OrdreBPS:= 301;
                  NbBPS:= 0;
                  IncBPS:= 3;
                  MaxBPS:= 310;
                  OrdreTSI:= 401;
                  NbTSI:= 0;
                  IncTSI:= 6;
                  MaxTSI:= 419;
                  OrdreExo:= 601;
                  NbExo:= 0;
                  IncExo:= 5;
                  MaxExo:= 646;
                  OrdrePrev:= 802;
                  NbPrev:= 0;
                  IncPrev:= 11;
                  MaxPrev:= 835;
                  end;

               While ((Segment<>'') and (Segment<>'S30.G01.00.001') and
                     (Segment<>'S41.G01.00.001')) do
                     begin
                     LL:= Length (S);
                     Virgule:= Pos(',', S);
                     Segment:= Copy (S,1,Virgule-1);
                     ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
                     TL:= TobLexique.FindFirst (['PDL_DADSSEGMENT'], [Segment],
                                                False);
{******************************************************************************}
                     if (Segment = 'S41.G01.00.003') then
                        begin
                        PerFin:= Copy (ValeurSeg, 3, 2)+Copy (ValeurSeg, 1, 2);
                        if (((DateDebMMJJ<=PerFin) and (PerFin<=DateFinMMJJ)) or
                           (PerFin<=DateDebMMJJ)) then
                           PerFin:= Copy (ValeurSeg, 1, 2)+'/'+
                                    Copy (ValeurSeg, 3, 2)+'/'+ExerciceDADS
                        else
                           PerFin:= Copy (ValeurSeg, 1, 2)+'/'+
                                    Copy (ValeurSeg, 3, 2)+'/'+
                                    IntToStr(StrToInt(ExerciceDads)-1);
                        end;

{££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££}
                     if ((Segment='S41.G01.00.012') or
                        (Segment='S41.G01.00.012.001')) then
                        begin
                        Q:= OpenSQL ('SELECT CO_LIBELLE FROM COMMUN WHERE'+
                                     ' CO_TYPE="PCT" AND'+
                                     ' CO_ABREGE="'+ValeurSeg+'"', True);
                        If Not Q.Eof then
                           ValeurSeg:= Q.FindField('CO_LIBELLE').AsString
                        else
                           ValeurSeg:= '';
                        Ferme(Q);
                        end;

                     if (Segment='S41.G01.00.013') then
                        begin
                        If (ValeurSeg = '90') then
                           ValeurSeg:= 'W'
                        else
                        If (ValeurSeg = '01') then
                           ValeurSeg:= 'C'
                        else
                        If (ValeurSeg = '02') then
                           ValeurSeg:= 'P'
                        else
                        If (ValeurSeg = '04') then
                           ValeurSeg:= 'I'
                        else
                        If (ValeurSeg = '05') then
                           ValeurSeg:= 'D'
                        else
                        If (ValeurSeg = '06') then
                           ValeurSeg:= 'S'
                        else
                        If (ValeurSeg = '07') then
                           ValeurSeg:= 'V'
                        else
                        If (ValeurSeg = '08') then
                           ValeurSeg:= 'O'
                        else
                        If (ValeurSeg = '09') then
                           ValeurSeg:= 'N'
                        else
                        If (ValeurSeg = '10') then
                           ValeurSeg:= 'F';
                        end;
{******************************************************************************}

                     if (Segment = 'S41.G01.00.019') then
                        begin
                        Libelle:= ValeurSeg;
                        TobUnSalarie.PutValueAllFille ('PDS_SALARIE', Libelle);
                        end
                     else
                     if (Libelle <> '') then
                        TobUnSalarieD.PutValue ('PDS_SALARIE', Libelle);

{££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££}
                     If (Segment='S41.G01.00.020') then
                        begin
                        If IsNumeric (ValeurSeg) then
                           ValeurSeg:= FloatToStr (StrToFloat(ValeurSeg)/100)
                        else
                           ValeurSeg:= '0';
                        end;

                     If (Segment = 'S41.G01.00.028') then
                        begin
                        If IsNumeric(ValeurSeg) then
                           ValeurSeg:= FloatToStr(StrToFloat(ValeurSeg)/100)
                        else
                           ValeurSeg:= '0';
                        end;

                     If (Segment='S41.G01.00.034') Then
                        begin
                        If ValeurSeg='01' Then
                           ValeurSeg:= 'F';
                        if ValeurSeg='02' Then
                           ValeurSeg:= 'E';
                        end;

                     If (Copy(Segment,1,10) = 'S41.G01.01') then
                        begin
                        If (Segment = 'S41.G01.01.001') then
                           begin
                           TL:= TobLexique.FindFirst(['PDL_DADSSEGMENT','ORDRESEG'],
                                                     [Segment,IntToStr(OrdreRC+NbRC*IncRC)+'L'],
                                                     False);
                           Segment:= Segment+'.'+IntToStr(OrdreRC+NbRC*IncRC);
                           If TL <> Nil then
                              T.PutValue(Segment+'L',
                                         RechDom (TL.GetValue('TABLETTE'),
                                                  ValeurSeg,False));
                           NbRC:= NbRC + 1;
                           end;

                        If (Segment='S41.G01.01.002') then
                           Segment:= Segment+'.'+
                                     IntToStr(OrdreRC+1+(NbRC-1)*IncRC);
                        end;

                     If (Copy(Segment,1,10) = 'S41.G01.02') then
                        begin
                        If (Segment = 'S41.G01.02.001') then
                           begin
                           TL:= TobLexique.FindFirst(['PDL_DADSSEGMENT','ORDRESEG'],
                                                     [Segment,IntToStr(OrdreBBS+NbBBS*IncBBS)+'L'],
                                                     False);
                           Segment:= Segment+'.'+IntToStr(OrdreBBS+NbBBS*IncBBS);
                           if TL <> Nil then
                              T.PutValue (Segment+'L',
                                          RechDom (TL.GetValue('TABLETTE'),
                                                   ValeurSeg, False));
                           NbBBS:= NbBBS + 1;
                           end;

                        if ((ExerciceDads>='2003') and
                           (Segment = 'S41.G01.02.002.001')) or
                           ((ExerciceDads<'2003') and
                           (Segment = 'S41.G01.02.002')) then
                           Segment:= Segment+'.'+
                                     IntToStr(OrdreBBS+1+(NbBBS-1)*IncBBS);

                        If (Segment = 'S41.G01.02.002.002') then
                           Segment:= Segment+'.'+
                                     IntToStr(OrdreBBS+1+(NbBBS-1)*IncBBS);
                        end;

                     If (Copy(Segment,1,10) = 'S41.G01.03') then
                        begin
                        If (Segment = 'S41.G01.03.001') then
                           begin
                           TL:= TobLexique.FindFirst(['PDL_DADSSEGMENT','ORDRESEG'],
                                                     [Segment,IntToStr(OrdreBPS+NbBPS*IncBPS)+'L'],
                                                     False);
                           Segment:= Segment+'.'+IntToStr(OrdreBPS+NbBPS*IncBPS);
                           If TL <> Nil then
                              T.PutValue (Segment+'L',
                                          RechDom (TL.GetValue('TABLETTE'),
                                                   ValeurSeg,False));
                           NbBPS:= NbBPS + 1;
                           end;

                        If ((ExerciceDads>='2003') and
                           (Segment = 'S41.G01.03.002.001')) or
                           ((ExerciceDads<'2003') and
                           (Segment = 'S41.G01.03.002')) then
                           Segment:= Segment+'.'+
                                     IntToStr(OrdreBPS+1+(NbBPS-1)*IncBPS);

                        If (Segment = 'S41.G01.03.002.002') then
                           Segment:= Segment+'.'+
                                     IntToStr(OrdreBPS+1+(NbBPS-1)*IncBPS);
                        end;

                     If (Copy(Segment,1,10) = 'S41.G01.04') then
                        begin
                        If (Segment = 'S41.G01.04.001') then
                           begin
                           TL:= TobLexique.FindFirst(['PDL_DADSSEGMENT','ORDRESEG'],
                                                     [Segment,IntToStr(OrdreTSI+NbTSI*IncTSI)+'L'],
                                                     False);
                           Segment:= Segment+'.'+IntToStr(OrdreTSI+NbTSI*IncTSI);
                           If TL <> Nil then
                              T.PutValue (Segment+'L',
                                          RechDom (TL.GetValue('TABLETTE'),
                                                   ValeurSeg,False));
                           NbTSI:= NbTSI + 1;
                           end;

                        If (Segment = 'S41.G01.04.002') then
                           Segment:= Segment+'.'+
                                     IntToStr(OrdreTSI+1+(NbTSI-1)*IncTSI);

                        If ((ExerciceDads>='2003') and
                           (Segment = 'S41.G01.04.003.001')) or
                           ((ExerciceDads<'2003') and
                           (Segment = 'S41.G01.04.003')) then
                           Segment:= Segment+'.'+
                                     IntToStr(OrdreTSI+2+(NbTSI-1)*IncTSI);

                        If (Segment = 'S41.G01.04.003.002') then
                           Segment:= Segment+'.'+
                                     IntToStr(OrdreTSI+2+(NbTSI-1)*IncTSI);

                        If ((ExerciceDads>='2003') and
                           (Segment = 'S41.G01.04.004.001')) or
                           ((ExerciceDads<'2003') and
                           (Segment = 'S41.G01.04.004')) then
                           Segment:= Segment+'.'+
                                     IntToStr(OrdreTSI+4+(NbTSI-1)*IncTSI);

                        If (Segment = 'S41.G01.04.004.002') then
                           Segment:= Segment+'.'+
                                     IntToStr(OrdreTSI+4+(NbTSI-1)*IncTSI);
                        end;

                     If (Copy(Segment,1,10) = 'S41.G01.06') then
                        begin
                        If (Segment = 'S41.G01.06.001') then
                           begin
                           TL:= TobLexique.FindFirst(['PDL_DADSSEGMENT','ORDRESEG'],
                                                     [Segment,IntToStr(OrdreExo+NbExo*IncExo)+'L'],
                                                     False);
                           Segment:= Segment+'.'+IntToStr(OrdreExo+NbExo*IncExo);
                           If TL <> nil then
                              T.PutValue (Segment+'L',
                                          RechDom (TL.GetValue('TABLETTE'),
                                                   ValeurSeg,False));
                           NbExo:= NbExo + 1;
                           end;

                        If ((ExerciceDads>='2003') and
                           (Segment = 'S41.G01.06.002.001')) or
                           ((ExerciceDads<'2003') and
                           (Segment = 'S41.G01.06.002')) then
                           Segment:= Segment+'.'+
                                     IntToStr(OrdreExo+1+(NbExo-1)*IncExo);

                        If (Segment = 'S41.G01.06.002.002') then
                           Segment:= Segment+'.'+
                                     IntToStr(OrdreExo+1+(NbExo-1)*IncExo);

                        If ((ExerciceDads>='2003') and
                           (Segment = 'S41.G01.06.003.001')) or
                           ((ExerciceDads<'2003') and
                           (Segment = 'S41.G01.06.003')) then
                           Segment:= Segment+'.'+
                                     IntToStr(OrdreExo+3+(NbExo-1)*IncExo);

                        If (Segment = 'S41.G01.06.003.002') then
                           Segment:= Segment+'.'+
                                     IntToStr(OrdreExo+3+(NbExo-1)*IncExo);
                        end;

                     if ((Segment='S41.G02.00.009') OR
                        (Segment='S41.G02.00.010')) then
                        ValeurSeg:= Copy(ValeurSeg, 2,1);

                     If (Segment='S44.G01.00.002') then
                        begin
                        If (IsNumeric (ValeurSeg)) then
                           ValeurSeg:= FloatToStr (StrToFloat(ValeurSeg)/100)
                        else
                           ValeurSeg:= '0';
                        end;

                     If (Copy(Segment,1,10) = 'S45.G01.01') then
                        begin
                        If (Segment = 'S45.G01.01.001') then
                           begin
                           Segment:= Segment+'.'+
                                     IntToStr(OrdrePrev+NbPrev*IncPrev);
                           NbPrev:= NbPrev + 1;
                           end;

                        If (Segment = 'S45.G01.01.002') then
                           Segment:= Segment+'.'+
                                     IntToStr(Ordreprev+1+(NbPrev-1)*IncPrev);

                        If (Segment = 'S45.G01.01.003') then
                           Segment:= Segment+'.'+
                                     IntToStr(Ordreprev+2+(NbPrev-1)*IncPrev);

                        If (Segment = 'S45.G01.01.004') then
                           Segment:= Segment+'.'+
                                     IntToStr(Ordreprev+3+(NbPrev-1)*IncPrev);

                        If (Segment = 'S45.G01.01.005') then
                           Segment:= Segment+'.'+
                                     IntToStr(Ordreprev+4+(NbPrev-1)*IncPrev);

                        If (Segment = 'S45.G01.01.006') then
                           Segment:= Segment+'.'+
                                     IntToStr(Ordreprev+5+(NbPrev-1)*IncPrev);

                        If (Segment = 'S45.G01.01.007.001') then
                           Segment:= Segment+'.'+
                                     IntToStr(Ordreprev+7+(NbPrev-1)*IncPrev);

                        If (Segment = 'S45.G01.01.008') then
                           Segment:= Segment+'.'+
                                     IntToStr(Ordreprev+8+(NbPrev-1)*IncPrev);

                        If (Segment = 'S45.G01.01.009') then
                           Segment:= Segment+'.'+
                                     IntToStr(Ordreprev+9+(NbPrev-1)*IncPrev);

                        If (Segment = 'S45.G01.01.010') then
                           Segment:= Segment+'.'+
                                     IntToStr(Ordreprev+10+(NbPrev-1)*IncPrev);
                        end;

                     If (Segment='S66.G01.00.003') then
                        begin
                        If (IsNumeric (ValeurSeg)) then
                           ValeurSeg:= FloatToStr (StrToFloat (ValeurSeg)/100)
                        else
                           ValeurSeg:= '0';
                        end;

                     If (Segment = 'S66.G01.00.009') then
                        begin
                        If (ValeurSeg = '01') then
                           ValeurSeg:= 'HEB'
                        else
                           ValeurSeg:= 'MEN';
                        end;

                     If (Segment='S66.G01.00.010') then
                        begin
                        If (IsNumeric (ValeurSeg)) then
                           ValeurSeg:= FloatToStr (StrToFloat (ValeurSeg)/100)
                        else
                           ValeurSeg:= '0';
                        end;

                     If (Segment='S66.G01.00.011') then
                        begin
                        If (IsNumeric (ValeurSeg)) then
                           ValeurSeg:= FloatToStr (StrToFloat (ValeurSeg)/100)
                        else
                           ValeurSeg:= '0';
                        end;

                     If (Segment = 'S66.G01.00.012') then
                        begin
                        Q:= OpenSQL ('SELECT CO_LIBELLE FROM COMMUN WHERE'+
                                     ' CO_TYPE="PSF" AND'+
                                     ' CO_ABREGE="'+ValeurSeg+'"', True);
                        If Not Q.Eof then
                           ValeurSeg:= Q.FindField('CO_LIBELLE').AsString
                        else
                           ValeurSeg:= '';
                        Ferme(Q);
                        end;

                     EnregistreSegment(TL, ExerciceDads, Segment, ValeurSeg, T);
{******************************************************************************}
                     TobUnSalarieD.PutValue ('PDS_TYPE', TypeD);
                     TobUnSalarieD.PutValue ('PDS_ORDRE', Ordre);
                     TobUnSalarieD.PutValue ('PDS_SEGMENT', Segment);
                     TobUnSalarieD.PutValue ('PDS_DONNEE', ValeurSeg);
                     TobUnSalarieD.PutValue ('PDS_EXERCICEDADS', ExerciceDads);
                     TobUnSalarieD.PutValue ('PDS_DATEDEBUT', IDate1900);
                     TobUnSalarieD.PutValue ('PDS_DATEFIN', IDate1900);
                     Readln (FLect,S);
                     LL:= Length (S);
                     Virgule:= Pos(',', S);
                     Segment:= Copy (S,1,Virgule-1);
                     Structure:= Copy (S,1,3);
                     ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
{££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££}
                     end;
               CreationTob:= False;
{******************************************************************************}
               end;

         TobUnSalarie.ParcoursTraitement (['PDS_ORDRE'], [Ordre], False,
                                          AffecteDebut);
         TobUnSalarie.ParcoursTraitement (['PDS_ORDRE'], [Ordre], False,
                                          AffecteFin);
         end;

      if ((Segment='S70.G01.00.001') or (Segment='S70.G01.00.002.001')) then
      //Si nouvel honoraire
         begin
         Ordre:= 0;

////////////////////////////////////////////////////////////////////////////////
         FreeAndNil (TobEntete);
         TobEntete:= Tob.Create('MonEntete',Nil,-1);
         T:= Tob.create('Filles',TobPeriodes,-1);
         T.AddchampSupValeur('DOSSIER', GetControlText('DOSSIER'), False);
         T.AddchampSupValeur('DATEDEB', GetControlText('DATEDEB'), False);
         T.AddchampSupValeur('DATEFIN', GetControlText('DATEFIN'), False);
         CreationSegments (ExerciceDads, T, TobEntete);
////////////////////////////////////////////////////////////////////////////////

         while (Structure = 'S70') do
               begin
               if ((Segment='S70.G01.00.001') or
                  (Segment='S70.G01.00.002.001')) then
                  begin
                  Ordre:= Ordre+1;
                  TobUnHonoraire:= Tob.Create ('DADSDETAIL', TobHonoraire, -1);
                  end;

               TobUnHonoraireD:= Tob.Create ('DADSDETAIL', TobUnHonoraire, -1);
               TobUnHonoraireD.PutValue ('PDS_SALARIE', '--H'+IntToStr (Ordre));
               TobUnHonoraireD.PutValue ('PDS_TYPE', TypeD);
               TobUnHonoraireD.PutValue ('PDS_ORDRE', Ordre);
               TobUnHonoraireD.PutValue ('PDS_SEGMENT', Segment);
               TobUnHonoraireD.PutValue ('PDS_DONNEE', ValeurSeg);
               TobUnHonoraireD.PutValue ('PDS_EXERCICEDADS', ExerciceDads);
               TobUnHonoraireD.PutValue ('PDS_DATEDEBUT', StrToDate (DateDeb));
               TobUnHonoraireD.PutValue ('PDS_DATEFIN', StrToDate (DateFin));
               Readln (FLect,S);
               LL:= Length (S);
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
               Structure:= Copy (S,1,3);
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
               end;
         end;

      if (Segment='S80.G01.00.001.001') then
      //Si nouvel établissement
         begin
         Ordre:= 0;

         while ((Structure = 'S80') or (Structure = 'S85')) do
               begin
               if (Segment='S80.G01.00.001.001') then
                  begin
                  Ordre:= Ordre+1;
                  TobUnEtab:= Tob.Create ('DADSDETAIL', TobEtabliss, -1);
                  end;

               TobUnEtabD:= Tob.Create ('DADSDETAIL', TobUnEtab, -1);
               TobUnEtabD.PutValue ('PDS_SALARIE', '**'+IntToStr (Ordre));
               TobUnEtabD.PutValue ('PDS_TYPE', TypeD);
               TobUnEtabD.PutValue ('PDS_ORDRE', Ordre);
               TobUnEtabD.PutValue ('PDS_EXERCICEDADS', ExerciceDads);
               TobUnEtabD.PutValue ('PDS_DATEDEBUT', StrToDate (DateDeb));
               TobUnEtabD.PutValue ('PDS_DATEFIN', StrToDate (DateFin));
               TobUnEtabD.PutValue ('PDS_SEGMENT', Segment);
               TobUnEtabD.PutValue ('PDS_DONNEE', ValeurSeg);

               Readln (FLect,S);
               LL:= Length (S);
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
               Structure:= Copy (S,1,3);
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
               end;
         end;

////////////////////////////////////////////////////////////////////////////////
      For i:=0 to TobEtabliss.Detail.Count-1 do
          begin
          NIC:= TobEtabliss.Detail[i].GetValue('NIC');
          T:= TobPeriodes.FindFirst(['S41.G01.00.005'], [NIC], False);
          while T<>nil do
                begin
                T.PutValue('S41.G01.00.005', TobEtabliss.Detail[i].GetValue('LIBELLE'));
                T:= TobPeriodes.FindNext(['S41.G01.00.005'], [NIC], False);
                end;
          end;
////////////////////////////////////////////////////////////////////////////////

      if (Segment = 'S90.G01.00.001') then
         begin
         Ordre:= 0;

         while (Structure = 'S90') do
               begin
               Readln (FLect,S);
               LL:= Length (S);
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
               Structure:= Copy (S,1,3);
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
               end;
         end;
      end;
CloseFile(FLect);

////////////////////////////////////////////////////////////////////////////////
Ordre:= 1;
For i := 0 to TobPeriodes.Detail.Count-1 do
    begin
    Salarie:= TobPeriodes.Detail[i].GetValue('S41.G01.00.019');
    TobPeriodes.Detail[i].AddchampSupValeur('PDE_SALARIE',Salarie,False);
    TobPeriodes.Detail[i].AddchampSupValeur('PDE_ORDRE',Ordre,False);
    Ordre:= Ordre+1;
    end;

TobPeriodes.Detail.Sort('DOSSIER;PDE_SALARIE;PDE_ORDRE');
FreeAndNil (TobLexique);
////////////////////////////////////////////////////////////////////////////////
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/07/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAYE;PGDADSU
*****************************************************************}
function TOF_PGEDIT_DADSIMPORT.AffecteDebut(T : TOB) : integer;
begin
T.PutValue ('PDS_DATEDEBUT', StrToDate (PerDeb));
result:= 0;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/07/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAYE;PGDADSU
*****************************************************************}
function TOF_PGEDIT_DADSIMPORT.AffecteFin(T : TOB) : integer;
begin
T.PutValue ('PDS_DATEFIN', StrToDate (PerFin));
result:= 0;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/08/2006
Modifié le ... :   /  /
Description .. : Création des champs correspondants aux segments de la
Suite ........ : DADS
Mots clefs ... : PAYE;PGDADSU
*****************************************************************}
procedure TOF_PGEDIT_DADSIMPORT.CreationSegments (ExerciceDads:string; var T:tob;
                                                  TobEntete:tob);
var
x : Integer;
Nature, OrdreSeg, Segment, Tablette, ValeurSeg : string;
begin
For x:=0 to TobLexique.Detail.Count-1 do
    begin
    Segment:= TobLexique.Detail[x].GetValue('PDL_DADSSEGMENT');
    Nature:= TobLexique.Detail[x].GetValue('PDL_DADSNATURE');
    OrdreSeg:= TobLexique.Detail[x].GetValue('ORDRESEG');
    Tablette:= TobLexique.Detail[x].GetValue('TABLETTE');
    If OrdreSeg <> '' then
       Segment:= Segment+'.'+OrdreSeg;
    If (TobEntete.FieldExists(Segment)) then
       ValeurSeg:= TobEntete.GetValue(Segment)
    else
       ValeurSeg:= '';

    If segment = 'S30.G01.00.007' then  // Civilité
       begin
       If ValeurSeg = '01' then
          ValeurSeg := 'Monsieur'
       else
       If ValeurSeg = '02' then
          ValeurSeg := 'Madame'
       else
       If ValeurSeg = '03' then
          ValeurSeg := 'Mademoiselle';
       end;

    If (Tablette <> '') and (ValeurSeg <> '') then
       //Recherhce tablette associé
       T.AddchampSupValeur(Segment+'L',RechDom(Tablette,ValeurSeg,False),False);
    If Nature = 'D' then  // Cas d'une date
       begin
       If Length(ValeurSeg) = 4 then
          ValeurSeg := ValeurSeg+ExerciceDads;
       If ValeurSeg <> '' then
          T.AddchampSupValeur (Segment, StrToDate (Copy(ValeurSeg,1,2)+'/'+
                                                   Copy(ValeurSeg,3,2)+'/'+
                                                   Copy(ValeurSeg,5,4)),False)
       else
          T.AddchampSupValeur(Segment,IDate1900,False);
       end
    else
    if nature = 'N' then //Cas d'un numérique
       begin
       If IsNumeric (ValeurSeg) then
          T.AddchampSupValeur (Segment, StrToFloat(ValeurSeg), False)
       else
          T.AddchampSupValeur(Segment,0,False);
       end
    else
          T.AddchampSupValeur(Segment,ValeurSeg,False);
    end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/07/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAYE;PGDADSU
*****************************************************************}
procedure TOF_PGEDIT_DADSIMPORT.ChargerLexiqueDads (Annee, NatureDADSU:String);
var
DadsValeur, Nature, Segment, StQ, StTypeDecla : string;
Q : TQuery;
i, Position, x : integer;
TL : Tob;
begin
StTypeDecla:= GetDADSUNature (NatureDADSU);
StQ:= 'SELECT *'+
      ' FROM DADSLEXIQUE WHERE'+
      ' PDL_EXERCICEDEB <= "'+Annee+'" AND'+
      ' (PDL_EXERCICEFIN >= "'+Annee+'" OR PDL_EXERCICEFIN = "")'+
      ' ORDER BY PDL_DADSSEGMENT';
Q:= OpenSQL (StQ, True);
TLexique.LoadDetailDB('DADSLEXIQUE','','',Q,False);
Ferme(Q);

StQ:= 'SELECT PDL_DADSSEGMENT, PDL_DADSNATURE, PDL_DADSVALEUR'+
      ' FROM DADSLEXIQUE WHERE'+
      ' PDL_DADSSEGMENT>"S30" AND'+
      ' PDL_DADSSEGMENT<"S50" AND'+
      ' PDL_EXERCICEDEB <= "'+Annee+'" AND'+
      ' (PDL_EXERCICEFIN >= "'+Annee+'" OR PDL_EXERCICEFIN = "")'+
      ' ORDER BY PDL_DADSSEGMENT';
Q:= OpenSQL (StQ, True);
TobLexique.LoadDetailDB('LesSegments','','',Q,False);
Ferme(Q);


For i := 0 to TobLexique.Detail.Count - 1 do
    begin
    Segment:= TobLexique.Detail[i].GetValue('PDL_DADSSEGMENT');
    Nature:= TobLexique.Detail[i].GetValue('PDL_DADSNATURE');
    TobLexique.Detail[i].AddchampSupValeur('ORDRESEG','',False);
    TobLexique.Detail[i].AddchampSupValeur('TABLETTE','',False);
    If Segment = 'S41.G01.01.001' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','101');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 4 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(101 + (x * 2)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       For x := 1 to 5 do
           begin
           TL := Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(99 + (x * 2))+'L',False);
           TL.AddchampSupValeur('TABLETTE','PGINSTITUTION',False);
           end;
       end
    else
    If (Segment='S41.G01.00.013') Then
       TobLexique.Detail[i].PutValue('TABLETTE','PGCONDEMPLOI')
    else
    If (Segment='S41.G01.00.034') Then
       TobLexique.Detail[i].PutValue('TABLETTE','PGTRAVAILETRANGER')
    else
    If Segment = 'S41.G01.01.002' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','102');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 4 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(102 + (x * 2)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    If Segment = 'S41.G01.02.001' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','201');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(201 + (x * 3)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       For x := 1 to 4 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(198 + (x * 3))+'L',False);
           TL.AddchampSupValeur('TABLETTE','PGBASEBRUTESPEC',False);
           end;
       end
    else
    If (Segment = 'S41.G01.02.002.001') or (Segment = 'S41.G01.02.002') then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','202');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(202 + (x * 3)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    If Segment = 'S41.G01.03.001' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','301');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(301 + (x * 3)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       For x := 1 to 4 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(298 + (x * 3))+'L',False);
           TL.AddchampSupValeur('TABLETTE','PGBASEPLAFSPEC',False);
           end;
       end
    else
    If (Segment = 'S41.G01.03.002.001') or (Segment = 'S41.G01.03.002') then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','302');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(302 + (x * 3)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    If Segment = 'S41.G01.04.001' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','401');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(401 + (x * 6)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       For x := 1 to 4 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(395 + (x * 6))+'L',False);
           TL.AddchampSupValeur('TABLETTE','PGSOMISO',False);
           end;
       end
    else
    If Segment = 'S41.G01.04.002' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','402');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(402 + (x * 6)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    If (Segment = 'S41.G01.04.003.001') or (Segment = 'S41.G01.04.003') then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','403');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(403 + (x * 6)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    If ((Segment='S41.G01.04.004.001') or (Segment='S41.G01.04.004.001')) then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','405');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(405 + (x * 6)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    If Segment = 'S41.G01.06.001' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','601');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 9 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(601 + (x * 5)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
        For x := 1 to 10 do
            begin
            TL:= Tob.Create('UneFille',TobLexique,-1);
            TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
            TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
            TL.AddchampSupValeur('ORDRESEG',IntToStr(596 + (x * 5))+'L',False);
            TL.AddchampSupValeur('TABLETTE','PGEXONERATION',False);
            end;
        end
    else
    If (Segment = 'S41.G01.06.002.001') or (Segment = 'S41.G01.06.002') then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','602');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 9 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(602 + (x * 5)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    If (Segment = 'S41.G01.06.003.001') or (Segment = 'S41.G01.06.003') then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','604');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 9 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(604 + (x * 5)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    If (Segment='S41.G02.00.009') Then
       TobLexique.Detail[i].PutValue('TABLETTE','PGCOLLEGEPRUD')
    else
    If (Segment='S41.G02.00.010') Then
       TobLexique.Detail[i].PutValue('TABLETTE','PGSECTIONPRUD')
    else
    If Segment = 'S45.G01.01.001' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','802');
       TobLexique.Detail[i].PutValue('TABLETTE','PGDADSPREV');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(802 + (x * 11)),False);
           TL.AddchampSupValeur('TABLETTE','PGDADSPREV',False);
           end;
       end
    else
    If Segment = 'S45.G01.01.002' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','803');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(803 + (x * 11)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    If Segment = 'S45.G01.01.003' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','804');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(804 + (x * 11)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    If Segment = 'S45.G01.01.004' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','805');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(805 + (x * 11)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    If Segment = 'S45.G01.01.005' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','806');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(806 + (x * 11)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    If Segment = 'S45.G01.01.006' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','807');
       TobLexique.Detail[i].PutValue('TABLETTE','PGDADSBASEPREV');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(807 + (x * 11)),False);
           TL.AddchampSupValeur('TABLETTE','PGDADSBASEPREV',False);
           end;
       end
    else
    If (Segment = 'S45.G01.01.007.001') or (Segment = 'S45.G01.01.007') then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','808');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(808 + (x * 11)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    If Segment = 'S45.G01.01.008' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','810');
       TobLexique.Detail[i].PutValue('TABLETTE','PGDADSPREVPOP');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(810 + (x * 11)),False);
           TL.AddchampSupValeur('TABLETTE','PGDADSPREVPOP',False);
           end;
       end
    else
    If Segment = 'S45.G01.01.009' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','811');
       TobLexique.Detail[i].PutValue('TABLETTE','PGSITUATIONFAMIL');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(811 + (x * 11)),False);
           TL.AddchampSupValeur('TABLETTE','PGSITUATIONFAMIL',False);
           end;
       end
    else
    If Segment = 'S45.G01.01.010' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','812');
       TobLexique.Detail[i].PutValue('TABLETTE','');
       For x := 1 to 3 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(812 + (x * 11)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    If (Segment='S66.G01.00.009') Then
       TobLexique.Detail[i].PutValue('TABLETTE','PGDADSBTPHORAIRE')
    else
    If Segment = 'S70.G01.01.001' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','101');
       TobLexique.Detail[i].PutValue('TABLETTE','PGDADSHONTYPEREM');
       For x := 1 to 4 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(101 + (x * 3)),False);
           TL.AddchampSupValeur('TABLETTE','PGDADSHONTYPEREM',False);
           end;
       end
    else
    If Segment = 'S70.G01.01.002' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','102');
       For x := 1 to 5 do
           begin
           TL:= Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(102 + (x * 3)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    if (TobLexique.Detail[i].GetValue('PDL_DADSVALEUR')<>'') then
       begin
       DadsValeur:= TobLexique.Detail[i].GetValue('PDL_DADSVALEUR');
       Position:= Pos ('T', DadsValeur);
       if (Position=1) then
          begin
          System.Delete (DadsValeur, 1, 2);
          Position:= Pos (';', DadsValeur);
          DadsValeur:= Copy (DadsValeur, 1, Position-1);
          TobLexique.Detail[i].PutValue('TABLETTE',DadsValeur);
          end;
       end
    else
    If (Segment='S41.G01.00.011') Then
       begin
       if (GetParamSocSecur('SO_PGPCS2003', 'X') = True) then
          TobLexique.Detail[i].PutValue('TABLETTE','PGCODEPCSESE')
       else
          TobLexique.Detail[i].PutValue('TABLETTE','PGCODEEMPLOI');
       end
    else
    If (Segment='S41.G01.00.017') Then
       TobLexique.Detail[i].PutValue('TABLETTE','PGCOEFFICIENT')
    else
    If (Segment='S66.G01.00.017') Then
       TobLexique.Detail[i].PutValue('TABLETTE','PGCOEFFICIENT')
    else
    If (Segment='S66.G01.00.022') Then
       TobLexique.Detail[i].PutValue('TABLETTE','PGDADSBTPAFFILIRC');
    end;
end;




{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/08/2006
Modifié le ... :   /  /
Description .. : Enregistrement de la valeur du segment à éditer dans la tob
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGEDIT_DADSIMPORT.EnregistreSegment (TL : tob; ExerciceDads,
                                                   Segment, ValeurSeg : string;
                                                   var T : tob);
var
Tablette, Nature : string;
begin
if (TL<>Nil) then
   begin
   Nature:= TL.GetValue('PDL_DADSNATURE');
   Tablette:= TL.GetValue('TABLETTE');
   If ((Segment = 'S41.G01.01.001.101') or (Segment = 'S41.G01.01.001.103') or
      (Segment = 'S41.G01.01.001.105') or (Segment = 'S41.G01.01.001.107') or
      (Segment = 'S41.G01.01.001.109') or (Segment = 'S41.G01.02.001.201') or
      (Segment = 'S41.G01.02.001.204') or (Segment = 'S41.G01.02.001.207') or
      (Segment = 'S41.G01.02.001.210') or (Segment = 'S41.G01.03.001.301') or
      (Segment = 'S41.G01.03.001.304') or (Segment = 'S41.G01.03.001.307') or
      (Segment = 'S41.G01.03.001.310') or (Segment = 'S41.G01.04.001.401') or
      (Segment = 'S41.G01.04.001.407') or (Segment = 'S41.G01.04.001.413') or
      (Segment = 'S41.G01.04.001.419') or (Segment = 'S41.G01.06.001.601') or
      (Segment = 'S41.G01.06.001.606') or (Segment = 'S41.G01.06.001.611') or
      (Segment = 'S41.G01.06.001.616') or (Segment = 'S41.G01.06.001.621') or
      (Segment = 'S41.G01.06.001.626') or (Segment = 'S41.G01.06.001.631') or
      (Segment = 'S41.G01.06.001.636') or (Segment = 'S41.G01.06.001.641') or
      (Segment = 'S41.G01.06.001.646')) then
      Tablette:= '';
   If (Tablette<>'') and (ValeurSeg<>'') then
      T.PutValue (Segment, RechDom (Tablette, ValeurSeg, False))
   else
      If ((Nature='D') or (segment='S41.G01.00.001') or
         (segment='S41.G01.00.003') or (segment='S46.G01.00.002') or
         (segment='S46.G01.00.003')) then
         begin
         If Length(ValeurSeg) = 4 then
            ValeurSeg:= ValeurSeg+ExerciceDads;
         If (ValeurSeg<>'') then
            T.PutValue (Segment, StrToDate (Copy (ValeurSeg,1,2)+'/'+
                                            Copy (ValeurSeg,3,2)+'/'+
                                            Copy (ValeurSeg,5,4)))
         else
            T.PutValue(Segment,IDate1900);
         end
      else
         if nature = 'N' then
            begin
            If IsNumeric (ValeurSeg) then
               T.PutValue(Segment,StrToFloat(ValeurSeg))
            else
               T.PutValue(Segment,0);
            end
         else
            T.PutValue(Segment,ValeurSeg);
   end
else
   begin
   Nature:= '';
   Tablette:= '';
   end;
end;

Initialization
  registerclasses ( [ TOF_PGEDIT_DADSIMPORT ] ) ;
end.

