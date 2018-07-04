{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/01/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGCONGESSPECENTRE ()
Mots clefs ... : TOF;PGCONGESSPECENTRE
*****************************************************************
PT1 13/10/2005 JL V_60  FQ 12227 Longueur enregistrement <> 300
PT2 18/07/2007 MF V_72  FQ 14588 : on encadre le nom du fichier (Chemin + Fichier) par des guillemets
}
Unit UTofPGCongesSpecEntre ;

Interface

Uses StdCtrls,Controls,Classes,Vierge,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,PG_OutilsEnvoi,UTob,
     PGOutils,PGOutils2,AGLInit,ParamSoc,ed_tools,ShellAPI,windows,EntPaie ;

Type
  TOF_PGCONGESSPECENTRE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    DateDebut,DateFin:TDateTime;
    FEcrt,FZ : TextFile;
    procedure ChargeZones ();
    procedure Generation (Sender : TObject);
    procedure GenereFichier();
    procedure EcritureLigne(Valeur,Format : String;Longueur : integer);
    Function ControleFichierCSP(Valeur : String;NumSegment : Integer) : String;
    Function ControlePonctuation(var Donnee : String) : String;
  end ;

Implementation

procedure TOF_PGCONGESSPECENTRE.OnArgument (S : String ) ;
var Arg : String;
    EtabPrinc,NumEntre : String;
    Q : TQuery;
begin
  Inherited ;
        Arg  :=  S;
        dateDebut := StrToDate(Trim(ReadTokenPipe(Arg,';')));
        DateFin := StrToDate(Trim(ReadTokenPipe(Arg,';'))) ;
        NumEntre := ReadTokenPipe(Arg,';');
        If NumEntre = '' then
        begin
                EtabPrinc := GetParamSoc('SO_ETABLISDEFAUT');
                Q := OpenSQL('SELECT ETB_ISNUMCPAY FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="'+EtabPrinc+'"',True);
                If Not Q.Eof Then NumEntre := Q.FindField('ETB_ISNUMCPAY').AsString;
                Ferme(Q);
        end;
        If NumEntre <> '' then
        begin
                SetControlText('NUMCOMPTEENTRE',Copy(NumEntre,1,8));
                SetControlText('CLECOMPTEENTRE',Copy(NumEntre,9,1));
        end;
        ChargeZones ();
        TFVierge(Ecran).BValider.OnClick  :=  Generation;
end ;

procedure TOF_PGCONGESSPECENTRE.ChargeZones ();
begin
        SetControlText('DATEDEBUT',DateToStr(DateDebut));
        SetControlText('DATEFIN',DateToStr(DateFin));
        SetControlText('PERIODICITE','T');
end;

procedure TOF_PGCONGESSPECENTRE.Generation (Sender : TObject);
begin
        try
                begintrans;
                GenereFichier();
                CommitTrans;
        Except
                Rollback;
                PGIBox ('Une erreur est survenue lors de la mise à jour de la base',Ecran.Caption);
        END;
end;

procedure TOF_PGCONGESSPECENTRE.GenereFichier();
var EnregEnvoi : TEnvoiSocial;
    QRechDelete : TQuery;
    StDelete,fichier,DateDebFichier,DateFinFichier,Libelle : String;
    Ordre,i,j : Integer;
    Q : TQuery;
    TobCongesSpec,T : Tob;
    FileN,FileZ : String;
    reponse : integer;
    NumSS,CleNumSS,CleNumConge,NumConge,NomNaiss,Prenom,NomEpoux,Pseudo : String;
    Emploi,Cadre,DatesTrav,NbJours,BaseConge,SalBrut,datePaiem,Lieu,DateDeliv : String;
    NomSign,Certificat,NumEntre,LettreCle,Libre,MessError : String;
    TabMess : array[1..22] of String;
    NbMess,VerifNumSS : Integer;
    Erreur : Boolean;
    DateDebSal,DateFinSal : TDateTime;
begin
        Erreur := False;
        DateDebFichier := FormatDateTime('ddmmyyyy',StrToDate(GetControlText('DATEDEBUT')));
        DateFinFichier := FormatDateTime('ddmmyyyy',StrTodate(GetControlText('DATEFIN')));
        fichier := 'CSP_'+ DateDebFichier+'au'+ DateFinFichier+'.CSP';
        Libelle := 'Congés spectacles du ' + getControlText('DATEDEBUT')+' au ' + GetControlText('DATEFIN');

        {$IFNDEF EAGLCLIENT}
        FileN  :=   V_PGI.DatPath+'\'+fichier;
        {$ELSE}
        FileN  :=   VH_Paie.PGCheminEagl+'\'+fichier;
        {$ENDIF}
        if FileExists(FileN) then
        begin
                reponse := HShowMessage('5;;Voulez-vous supprimer votre fichier Congés spectacles'+ FileN+';Q;YN;Y;N','','');
                if reponse = 6 then DeleteFile(PChar(FileN))
                else exit;
        end;
        AssignFile(FEcrt, FileN);
        ReWrite(FEcrt);
        Q := OpenSQL('SELECT * FROM CONGESSPEC WHERE'+
        ' PCS_DATEDEBUT>="'+UsDatetime(Datedebut)+'"'+
        ' AND PCS_DATEFIN<="'+UsDatetime(DateFin)+'"',True);
        TobCongesSpec := Tob.create('CONGESSPEC',Nil,-1);
        TobCongesSpec.LoadDetailDB('CONGESSPEC','','',Q,False);
        Ferme(Q);
        For i := 0 to TobCongesSpec.Detail.Count - 1 do
        begin
                TobCongesSpec.Detail[i].PutValue('PCS_NUMENTREP',GetControlText('NUMCOMPTEENTRE'));
                TobCongesSpec.Detail[i].PutValue('PCS_LETTRECLE',GetControlText('CLECOMPTEENTRE'));
                TobCongesSpec.Detail[i].UpdateDB(False);
        end;
        For i := 1 to 22 do TabMess[i] := '';
        {$IFNDEF EAGLCLIENT}
        if V_PGI.DatPath <>'' then FileZ:=V_PGI.DatPath+'\ErreurCSP.log'
        else FileZ:='C:\ErreurCSP.log';
        {$ELSE}
        if VH_Paie.PGCheminEagl <>'' then FileZ:=VH_Paie.PGCheminEagl+'\ErreurCSP.log'
        else FileZ:='C:\ErreurCSP.log';
        {$ENDIF}
        if SupprimeFichier(FileZ)=False then exit;
        AssignFile(FZ, FileZ);
        {$i-} ReWrite(FZ); {$i+}
        If IoResult<>0 Then Begin PGIBox('Fichier inaccessible : '+FileZ,'Abandon du traitement');Exit ; End;
        writeln(FZ,'Création fichier Congés spectacles : Gestion des messages d''erreur.');
        try
                begintrans;
                InitMoveProgressForm (NIL,'Début du traitement', 'Veuillez patienter SVP ...',TobCongesSpec.Detail.Count - 1,FALSE,TRUE);
                For i:= 0 to TobCongesSpec.Detail.Count - 1 do
                begin
                        NbMess := 0;
                        NumSS := TobCongesSpec.Detail[i].GetValue('PCS_NUMEROSS');
                        MessError := ControleFichierCSP(NumSS,1);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        CleNumSS := TobCongesSpec.Detail[i].GetValue('PCS_CLENUMSS');
                        MessError := ControleFichierCSP(CleNumSS,2);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        CleNumConge := PGUpperCase(TobCongesSpec.Detail[i].GetValue('PCS_CLENUMCONG'));
                        MessError := ControleFichierCSP(CleNumConge,3);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        NumConge := TobCongesSpec.Detail[i].GetValue('PCS_NUMCONGES');
                        MessError := ControleFichierCSP(NumConge,4);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        NomNaiss := PGUpperCase(TobCongesSpec.Detail[i].GetValue('PCS_NOMNAISS'));
                        Nomnaiss := ControlePonctuation(NomNaiss);
                        MessError := ControleFichierCSP(NomNaiss,5);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        Prenom := PGUpperCase(TobCongesSpec.Detail[i].GetValue('PCS_PRENOM'));
                        Prenom := ControlePonctuation(Prenom);
                        MessError := ControleFichierCSP(Prenom,6);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        NomEpoux := PGUpperCase(TobCongesSpec.Detail[i].GetValue('PCS_NOMEPOUX'));
                        NomEpoux := ControlePonctuation(NomEpoux);
                        MessError := ControleFichierCSP(NomEpoux,7);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        Pseudo := PGUpperCase(TobCongesSpec.Detail[i].GetValue('PCS_PSEUDO'));
                        Pseudo := ControlePonctuation(Pseudo);
                        MessError := ControleFichierCSP(Pseudo,8);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        Emploi := PGUpperCase(Rechdom('PGLIBEMPLOI',TobCongesSpec.Detail[i].GetValue('PCS_EMPLOI'),False));
                        If length(emploi) > 15 then Emploi := Copy(Emploi,1,15);
                        MessError := ControleFichierCSP(Emploi,9);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        Cadre := PGUpperCase(TobCongesSpec.Detail[i].GetValue('PCS_CADRE'));
                        MessError := ControleFichierCSP(Cadre,10);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        NbJours := TobCongesSpec.Detail[i].GetValue('PCS_NBJOURS');
                        MessError := ControleFichierCSP(NbJours,12);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        BaseConge := TobCongesSpec.Detail[i].GetValue('PCS_BASECONGE');
                        MessError := ControleFichierCSP(BaseConge,13);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        SalBrut := TobCongesSpec.Detail[i].GetValue('PCS_SALBRUT');
                        MessError := ControleFichierCSP(SalBrut,14);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        datePaiem := TobCongesSpec.Detail[i].GetValue('PCS_DATEPAIEM');
                        MessError := ControleFichierCSP(datePaiem,15);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        Lieu := PGUpperCase(TobCongesSpec.Detail[i].GetValue('PCS_LIEU'));
                        MessError := ControleFichierCSP(Lieu,16);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        DateDeliv := TobCongesSpec.Detail[i].GetValue('PCS_DATEDELIV');
                        MessError := ControleFichierCSP(DateDeliv,17);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        NomSign := PGUpperCase(TobCongesSpec.Detail[i].GetValue('PCS_NOMSIGN'));
                        MessError := ControleFichierCSP(NomSign,18);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        Certificat := TobCongesSpec.Detail[i].GetValue('PCS_NUMCERTIFICAT');
                        MessError := ControleFichierCSP(Certificat,19);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        NumEntre := TobCongesSpec.Detail[i].GetValue('PCS_NUMENTREP');
                        MessError := ControleFichierCSP(NumEntre,20);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        LettreCle := PGUpperCase(TobCongesSpec.Detail[i].GetValue('PCS_LETTRECLE'));
                        MessError := ControleFichierCSP(LettreCle,21);
                        If MessError <> '' then
                        begin
                                NbMess := NbMess + 1;
                                TabMess[NbMess] := MessError;
                        end;
                        DatesTrav := TobCongesSpec.Detail[i].GetValue('PCS_DATESTRAVAIL');
                        EcritureLigne(NumSS,'I',13);
                        EcritureLigne(CleNumSS,'I',2);
                        EcritureLigne(CleNumConge,'S',1);
                        EcritureLigne(NumConge,'I',7);
                        EcritureLigne(NomNaiss,'S',20);
                        EcritureLigne(Prenom,'S',12);
                        EcritureLigne(NomEpoux,'S',20);
                        EcritureLigne(Pseudo,'S',20);
                        EcritureLigne(Emploi,'S',15);
                        EcritureLigne(Cadre,'S',1);
                        EcritureLigne(DatesTrav,'I',16);
                        EcritureLigne(NbJours,'I',3);
                        EcritureLigne(BaseConge,'I',9);
                        EcritureLigne(SalBrut,'I',9);
                        EcritureLigne(datePaiem,'I',8);
                        EcritureLigne(Lieu,'S',20);
                        EcritureLigne(DateDeliv,'I',8);
                        EcritureLigne(NomSign,'S',20);
                        EcritureLigne(Certificat,'I',8);
                        EcritureLigne(NumEntre,'I',8);
                        EcritureLigne(LettreCle,'S',1);
                        Libre := '';
                        DateDebSal := TobCongesSpec.Detail[i].GetValue('PCS_DATEDEBUT');
                        DateFinSal := TobCongesSpec.Detail[i].GetValue('PCS_DATEFIN');
                        If NbJours = '' then Libre := 'JOURS A ZERO'
                        else
                                If StrToInt(NbJours) = 0 then Libre :='JOURS A ZERO'
                                else
                                begin
                                        If ((PLUSDATE (DateDebSal,Trunc(StrToInt(NbJours)-1),'J')) < DateFinSal) then Libre := 'JOURS> A ECART DE DATE';
                                end;
                        EcritureLigne(Libre,'S',60);
                        EcritureLigne('','S',19); //PT1
                        TobCongesSpec.Detail[i].PutValue('PCS_LIBRE',Libre);
                        TobCongesSpec.Detail[i].UpdateDB(False);
                        if NbMess >= 1 then
                        begin
                                Writeln (FZ,'******************* Salarié '+NomNaiss+' '+Prenom+
                                ', période du'+dateToStr(DateDebSal)+
                                'au'+DateToStr(DateFinSal));
                                Erreur := True;
                        end;
                        For j := 1 to NbMess do
                        begin
                                writeln (FZ,TabMess[j]);
                        end;
                        WriteLN(FEcrt,'');
                        MoveCurProgressForm(NomNaiss+' '+Prenom);
                end;
                FiniMoveProgressForm;
                CloseFile (FZ);
                CloseFile(FEcrt);
                If Erreur = True then
                begin
                        DeleteFile(PChar(FileN));
                        i := PGIAsk('Il y a des erreurs, #13#10 Voulez vous consulter maintenant le fichier '+FileZ,Ecran.Caption);
                        If i = MrYes then
// d PT2
                        begin
                        {$IFDEF EAGLCLIENT}
                          FileN  :=  '"'+ VH_Paie.PGCheminEagl+'\'+fichier+'"';
                        {$ENDIF}
                        ShellExecute( 0, PCHAR('open'),PChar('WordPad'), PChar(FileZ),Nil,SW_RESTORE);
                        end;
// f PT2
                end
                Else
                begin
                StDelete  :=  'SELECT MAX(PES_CHRONOMESS) AS MAXI FROM ENVOISOCIAL';
                QRechDelete := OpenSQL(StDelete,TRUE) ;
                if Not QRechDelete.EOF then Ordre  :=  QRechDelete.FindField ('MAXI').AsInteger+1
                else Ordre  :=  1;
                Ferme(QRechDelete);
                ChargeTOBENVOI ;
                EnregEnvoi.Ordre  :=  Ordre;
                EnregEnvoi.TypeE  :=  'PCS';
                EnregEnvoi.Millesime  :=  Copy(GetControlText('DATEFIN'),7,4);;
                EnregEnvoi.Periodicite  :=  'T';
                EnregEnvoi.DateD  :=  StrToDate(GetControltext('DATEDEBUT'));
                EnregEnvoi.DateF  :=  StrToDate(GetControltext('DATEFIN'));
                EnregEnvoi.Siret  :=  '';
                EnregEnvoi.Fraction  :=  '';
                EnregEnvoi.Libelle  :=  Libelle;
                EnregEnvoi.Size  :=  1;
                EnregEnvoi.NomFic  :=  fichier;
                EnregEnvoi.Statut  :=  '005';
                EnregEnvoi.Monnaie  :=  'EUR';
                EnregEnvoi.Inst := 'ZCSP';
                CreeEnvoi (EnregEnvoi);
                LibereTOBENVOI;
                        PGIBox('Le fichier est prêt à être envoyé,#13#10 la prochaine étape est la gestion des envois (module paramètres)',Ecran.caption);
                end;
                TobCongesSpec.Free;
                CommitTrans;
        Except
                Rollback;
                CloseFile (FZ);
                CloseFile(FEcrt);
                If TobCongesSpec <> Nil then TobCongesSpec.Free;
        end;
end;

procedure TOF_PGCONGESSPECENTRE.EcritureLigne(Valeur,Format : String;Longueur : integer);
var i : Integer;
begin
        If Format = 'I' Then
        begin
                If Length(Valeur)<Longueur Then
                begin
                        For i := Length(Valeur) to Longueur-1 do
                        begin
                                Valeur := '0'+Valeur;
                        end;
                end;
        end;
        If Format = 'S' Then
        begin
                If Length(Valeur)<Longueur Then
                begin
                        For i := Length(Valeur) to Longueur-1 do
                        begin
                                Valeur := Valeur+' ';
                        end;
                end;
        end;
        Write(FEcrt,Valeur);
end;

Function TOF_PGCONGESSPECENTRE.ControleFichierCSP(Valeur : String;NumSegment : Integer) : String;
begin
        Result := '';
        Case NumSegment of
                1 : //N° Sécu
                begin
                end;
                2 :  // Clé N° Sécu
                begin
                end;
                3 : //CléNum congé spectacle
                begin
                end;
                4 : // Num Congé Spectacle
                begin
                end;
                5 : //Nom de naissance
                begin
                        If Valeur ='' then Result := 'Le nom de naissance n''est pas renseigné et est obligatoire';
                end;
                6 :  //Prénom
                begin
                        If Valeur ='' then Result := 'Le prénom n''est pas renseigné et est obligatoire';
                end;
                7 : // Nom époux
                begin
                end;
                8 : // Pseudo
                begin
                end;
                9 : //EMPLOI
                begin
                        If Valeur ='' then Result := 'L''emploi n''est pas renseigné et est obligatoire';
                end;
                10 :  //Cadre
                begin
                        If Valeur ='' then Result := 'La zone cadre/non cadre n''est pas renseigné et est obligatoire';
                end;
                11 : // Dates travail
                begin
                        If Valeur ='' then Result := 'Les dates de tavail ne sont pas renseignées et est obligatoire';
                end;
                12 : //Nb Jours ou cachet
                begin
                        If Valeur ='' then Result := 'Le nombre de jours ou de cachets n''est pas renseigné et est obligatoire';
                end;
                13 : // BASE CONGES
                begin
                        If Valeur ='' then Result := 'La base congés n''est pas renseignée et est obligatoire';
                        If Valeur ='0' then Result := 'La base congés est égal à 0';
                end;
                14 :   // SALAIRE BRUT
                begin
                        If Valeur ='' then Result := 'Le salaire brut n''est pas renseigné et est obligatoire';
                        If Valeur ='0' then Result := 'Le salaire brut est égal à 0';
                end;
                15 : // Date paiement
                begin
                        If Valeur ='' then Result := 'La date de paiement n''est pas renseignée et est obligatoire';
                end;
                16 : // Lieu
                begin
                        If Valeur ='' then Result := 'Le lieu n''est pas renseigné et est obligatoire';
                end;
                17 : // Date délivrance
                begin
                        If Valeur ='' then Result := 'La date de délivrance n''est pas renseignée et est obligatoire';
                end;
                18 : // Signataire
                begin
                        If Valeur ='' then Result := 'Le nom du signataire n''est pas renseigné et est obligatoire';
                end;
                19 : // N° Certificat
                begin
                        If Valeur ='' then Result := 'Le n° de certificat n''est pas renseigné et est obligatoire';
                end;
                20 : // Num compte entreprise
                begin
                        If Valeur ='' then Result := 'Le numéro de compte entreprise n''est pas renseigné et est obligatoire';
                end;
                21 : // Lettre clé num compte entreprise
                begin
                        If Valeur ='' then Result := 'La lettre clé du numéro de compte entreprise n''est pas renseignée et est obligatoire';
                end;
                22 : // Libre
                begin
                end;
        end;
end;

Function TOF_PGCONGESSPECENTRE.ControlePonctuation(var Donnee : String) : String;
var Longueur,i : Integer;
    Ch : Char;
begin
        Longueur := Length(Donnee);
        For i := 1 to Longueur do
        begin
                Ch := Donnee[i];
                if Ch in [' ',',','-','.','/','`','¨','°','´'] then Ch := ' ';
                Donnee[i] := Ch;
        end;
        Result := Donnee;
end;

Initialization
  registerclasses ( [ TOF_PGCONGESSPECENTRE ] ) ;
end.

