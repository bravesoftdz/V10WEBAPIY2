{***********UNITE*************************************************
Auteur  ...... :  EPI
Créé le ...... : 20/03/2006
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PROCESSUS (PROCESSUS)
Mots clefs ... : TOM;PROCESSUS
*****************************************************************}
{
PT1 14/04/2006 V_65 EPI FQ 12791 gestion du détail
PT2 27/12/2006 V_80 FCO FQ 13253 supprimer le détail si le processus n'est finalement pas créé
PT3 26/03/2007 V_80 GGU FQ 13973 Perte du détail en PCL pour les processus STD ou CEG à cause d'une mauvaise valeur du nodossier
PT4 13/09/2007 V_80 FCO FQ 14388 Concepts
}
Unit PROCESSUS_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}  Fiche,  FichList, FE_Main,
{$else}
     eFiche, eFichList, MainEAGL,
{$ENDIF}
     forms,  sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox,UTOM, UTob ,
     HTB97,
     EntPaie;

Type
  TOM_PROCESSUS = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  private
    LectureSeule, CEG, STD, DOS : boolean;
    // procedure BTnInsClick(Sender: TObject);
    procedure TestValidite(var LastError : integer;var LastErrorMsg : WideString);  //PT3
    procedure BTnDetClick(Sender: TObject);
  end ;

Implementation
uses PgOutils2, Pgoutils;

procedure TOM_PROCESSUS.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_PROCESSUS.OnDeleteRecord ;
var
  St : String;
begin
   Inherited ;
   St := GetField('PRO_CODEPROC');
   ExecuteSQL('DELETE FROM PROCESSUSDET WHERE PRD_CODEPROC = "'+St+'"');
end ;

procedure TOM_PROCESSUS.OnUpdateRecord ;
var
  CodeProc,Ordre : String;
  Trouve : Boolean;
  Q : TQuery;
begin
  Inherited ;

  TestValidite(LastError,LastErrorMsg);
  // PT1 Ajout controle sur la présence de lignes détail
  Trouve := False;
 	CodeProc := GetField('PRO_CODEPROC');
	Ordre := 'SELECT PRD_CODEPROC FROM PROCESSUSDET WHERE PRD_CODEPROC="';
 	Ordre := Ordre + CodeProc + '"';
  Q:= OpenSQL(Ordre,TRUE) ;
 	While not Q.EOF do
   	begin
      Trouve := True;
  	  Q.Next ;
	  end ;
    Ferme(Q) ;
  if (Trouve = False) then
  begin
    LastError := 1;
    LastErrorMsg := 'Le détail du processus n''est pas saisi';
    SetFocusControl('PRO_PROCESTYPE');
    exit;
  end;
end ;

procedure TOM_PROCESSUS.OnLoadRecord ;
begin
  Inherited ;
  AccesPredefini('TOUS', CEG, STD, DOS);
  LectureSeule := FALSE;
  if (Getfield('PRO_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
  end
  else
    if (Getfield('PRO_PREDEFINI') = 'STD') then
    begin
      LectureSeule := (STD = False);
      PaieLectureSeule(TFFiche(Ecran), (STD = False));
    end
    else
      if (Getfield('PRO_PREDEFINI') = 'DOS') then
      begin
        LectureSeule := (DOS = False); //PT4
        PaieLectureSeule(TFFiche(Ecran), (DOS = False));   //PT4
      end;
    SetControlEnabled('PRO_PREDEFINI', False);
    SetControlEnabled('PRO_NODOSSIER', False);
    SetControlEnabled('PRO_CODEPROC', False);
    SetControlEnabled('PRO_PROCESTYPE', False);
//PT4    SetControlEnabled('PRO_PROSELECT', True);
    // Appel fiche en insertion
    if DS.State in [dsInsert] then
    begin
      LectureSeule := FALSE;
      PaieLectureSeule(TFFiche(Ecran), False);
      SetControlEnabled('PRO_PREDEFINI', True);
      SetField('PRO_PREDEFINI', 'DOS');
      // PT1 SetControlVisible('BDetail',False);
      // SetControlVisible('BTrait',False);
  end;
end ;

procedure TOM_PROCESSUS.OnChangeField ( F: TField ) ;
var
  Pred : String;

begin
  Inherited ;
  // Controle code prédéfini
  if (F.FieldName = 'PRO_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Pred := GetField('PRO_PREDEFINI');
    if Pred = '' then exit;
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de rubrique prédéfinie CEGID', 'Accès refusé');
      Pred := 'DOS';
      SetControlProperty('PRO_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de rubrique prédéfinie Standard', 'Accès refusé'); // PT28
      Pred := 'DOS';
      SetControlProperty('PRO_PREDEFINI', 'Value', Pred);
    end;
  end;
end ;

procedure TOM_PROCESSUS.OnArgument ( S: String ) ;
var
  //BtnIns,BtnDet,BtnTrt : TToolBarButton97;
  BtnDet : TToolBarButton97;
begin
  Inherited ;
  // BtnIns := TToolBarButton97(GetControl('BINSERT'));
  BtnDet := TToolBarButton97(GetControl('BDETAIL'));
  // BtnTrt := TToolBarButton97(GetControl('BTRAIT'));
  // if BtnIns <> nil then BtnIns.OnClick := BTnInsClick;
  if BtnDet <> nil then BtnDet.OnClick := BTnDetClick;
  //if BtnTrt <> nil then BtnTrt.OnClick := BTnTrtClick;
end ;

procedure TOM_PROCESSUS.OnClose ;
var
  //Début PT2
  St      : String;
  Ordre   : String;
  Q, QDet : TQuery;
  //Fin PT2
begin
  Inherited ;
  //Début PT2
  St    := GetField('PRO_CODEPROC');
	Ordre := 'SELECT PRO_CODEPROC FROM PROCESSUS WHERE PRO_CODEPROC="' + St + '"';
  Q     := OpenSQL(Ordre,TRUE) ;
  // Si le processus n'existe pas, vérifier qu'il n'y a pas de détail existant
  if Q.Eof then
  begin
  	Ordre := 'SELECT PRD_CODEPROC FROM PROCESSUSDET WHERE PRD_CODEPROC="' + St + '"';
    QDet  := OpenSQL(Ordre,TRUE) ;
    if not QDet.Eof then
    begin
      ExecuteSQL('DELETE FROM PROCESSUSDET WHERE PRD_CODEPROC = "' + St + '"');
    end;
    Ferme(QDet) ;
  end;
  Ferme(Q) ;
  //Fin PT2
end ;

procedure TOM_PROCESSUS.OnCancelRecord ;
begin
  Inherited ;
end ;

//  Insertion d'une ligne
{procedure TOM_PROCESSUS.BTnInsClick(Sender: TObject);
begin
  SetField('PRO_PREDEFINI', 'DOS');
  SetControlEnabled('PRO_PREDEFINI',False);
  SetControlVisible('BDetail',False);
  // SetControlVisible('BTrait',False);
end;}

// Affichage du détail d'un processus
procedure TOM_PROCESSUS.BTnDetClick(Sender: TObject);
var
  St : String;
  // PT1 début
  Trouve : Boolean;
  Ordre : String;
  Q : TQuery;
  // PT1 fin

begin
  // PT1 ajout libellé sur entete écran
  // PT1 St :=   GetField('PRO_PREDEFINI') + ';' + GetField('PRO_NODOSSIER') + ';' +
  // PT1        GetField('PRO_CODEPROC') + ';' + GetField('PRO_PROCESTYPE')+ ';';
  TestValidite(LastError,LastErrorMsg); //PT3
  if LastError = 0 then
  begin
    St :=   GetField('PRO_PREDEFINI') + ';' + GetField('PRO_NODOSSIER') + ';' +
            GetField('PRO_CODEPROC') + ';' + GetField('PRO_PROCESTYPE')+ ';' +
            GetField('PRO_LIBELLEPROC') + ';';
    AGLLanceFiche('PAY','PROCESSUSDET_MUL','','',St);
    // PT1 début controle présence lignes détail
    Trouve := False;
    Ordre := 'SELECT PRD_CODEPROC FROM PROCESSUSDET WHERE PRD_CODEPROC="';
    Ordre := Ordre + GetField('PRO_CODEPROC') + '"';
    Q:= OpenSQL(Ordre,TRUE) ;
    While not Q.EOF do
      begin
        Trouve := True;
        Q.Next ;
      end ;
      Ferme(Q) ;
    if (Trouve = False) then
    begin
      PgiBox ('','Le détail du processus a été supprimé');
    end;
    // PT1 fin
  end else begin
    PGIError(TraduireMemoire(LastErrorMsg),Ecran.Caption);
  end;
end;

procedure TOM_PROCESSUS.TestValidite(var LastError : integer;var LastErrorMsg : WideString);
var
  Boucle,Predef,Typent,CodeProc,Ordre : String;
  Trouve : Boolean;
  Q : TQuery;
begin
  // Debut PT3  (mise en fonction des test de validité)
  SetFocusControl('PRO_LIBELLEPROC');
  if (DS.State = dsinsert) then
  begin
    Predef := GetField('PRO_PREDEFINI');
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Predef = 'CEG') and (CEG = FALSE) then
    begin
      Predef := 'DOS';
      LastError := 1;
      //LastErrorMsg := 'Vous ne pouvez pas créer de rubrique prédéfinie CEGID';
      SetControlProperty('PRO_PREDEFINI', 'Value', Predef);
      SetFocusControl('PRO_PREDEFINI');
    exit;
    end;
    if (Predef = 'STD') and (STD = FALSE) then
    begin
      Predef := 'DOS';
      LastError := 1;
      //LastErrorMsg := 'Vous ne pouvez pas créer de rubrique prédéfinie Standard';
      SetControlProperty('PRO_PREDEFINI', 'Value', Predef);
      SetFocusControl('PRO_PREDEFINI');
    end;
  end;

  if (DS.State = dsinsert) then
  begin
    if (GetField('PRO_PREDEFINI') <> 'DOS') then
      SetField('PRO_NODOSSIER', '000000')
    else
      SetField('PRO_NODOSSIER', PgRendNoDossier());
  end;

  if (GetField('PRO_CODEPROC') = '') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner le code du processus';
    SetFocusControl('PRO_CODEPROC');
    exit;
  end;

  TypEnt := GetField('PRO_PROCESTYPE');
  if (TypEnt = '') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner le champ type de processus';
    SetFocusControl('PRO_PROCESTYPE');
    exit;
  end;

  if (DS.State = dsinsert) then
  begin
	  Trouve := False;
  	CodeProc := GetField('PRO_CODEPROC');
	  Ordre := 'SELECT PRO_CODEPROC FROM PROCESSUS WHERE PRO_CODEPROC="';
  	Ordre := Ordre + CodeProc + '"';
	  Q:= OpenSQL(Ordre,TRUE) ;
  	While not Q.EOF do
    	begin
	    Trouve := True;
  	  Q.Next ;
	    end ;
  	  Ferme(Q) ;

	    If Trouve = True then
  	  begin
    	  LastError := 1;
      	LastErrorMsg := 'Ce processus existe déjà';
	      SetFocusControl('PRO_PROCESTYPE');
  	    exit;
    	end;
    end;

  Boucle := GetField('PRO_PROSELECT');
  if (Boucle = '') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez choisir un traitement';
    SetFocusControl('PRO_PROSELECT');
    exit;
  end;
  //Fin PT3
end;

Initialization
  registerclasses ( [ TOM_PROCESSUS ] ) ;
end.
