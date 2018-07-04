{***********UNITE*************************************************
Auteur  ...... : EPI
Créé le ...... : 21/03/2006
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PROCESSUSDET (PROCESSUSDET)
Mots clefs ... : TOM;PROCESSUSDET
*****************************************************************}
{
PT01 27/04/2006 V_65 EPI FQ 12791 ajout apostrophes sur les messages
PT2  13/09/2007 V_80 FCO FQ 14388 Concepts
}
Unit PROCESSUSDET_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList,
{$else}
     eFiche,
     eFichList,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     UTob ,
     HTB97;

Type
  TOM_PROCESSUSDET = Class (TOM)
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
    Protype,Arg,predef,dossier,code      : string;
    procedure BTnInsClick(Sender: TObject);
    end ;

Implementation
uses PgOutils2, Pgoutils;

procedure TOM_PROCESSUSDET.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_PROCESSUSDET.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_PROCESSUSDET.OnUpdateRecord ;
var
  Ordre,CodeProc,Trt : String;
  Trouve : boolean;
  Q : TQuery;
  No,Num : Integer;
begin
  Inherited ;
  Trouve := False;
 	CodeProc := GetField('PRD_CODEPROC');
 	Trt := GetField('PRD_PROCESACTION');
  No := StrToInt(Getfield('PRD_NOORDRE'));

 	If No < 1 then
	  begin
 		  LastError := 1;
   		LastErrorMsg := 'Le numéro d''ordre doit être supérieur à 0 ';
	    SetFocusControl('PRD_NOORDRE');
  	  exit;
	 	end;

  // Recherche unicité du numéro ordre
  if (DS.State = dsinsert) then
  begin
	  Ordre := 'SELECT PRD_NOORDRE FROM PROCESSUSDET WHERE PRD_CODEPROC="';
 		Ordre := Ordre + CodeProc + '"' +'AND PRD_NOORDRE="';
	  Ordre := Ordre + INttostr(No) + '"';
  	Q:= OpenSQL(Ordre,TRUE) ;
	  if not q.eof then
  	begin
	    Trouve := True;
  	end;
	  Ferme(Q) ;
  	If Trouve = True then
	  begin
 		  LastError := 1;
   		LastErrorMsg := 'Le numéro d''ordre doit être unique ';
	    SetFocusControl('PRD_NOORDRE');
  	  exit;
	 	end;
  end;

  // Recherche unicité du traitement
  Ordre := 'SELECT PRD_CODEPROC,PRD_NOORDRE FROM PROCESSUSDET WHERE PRD_CODEPROC="';
 	Ordre := Ordre + CodeProc + '"' +'AND PRD_PROCESACTION="';
  Ordre := Ordre + Trt + '"';
  Q:= OpenSQL(Ordre,TRUE) ;
  if not q.eof then
  begin
    Trouve := True;
    num := Q.Fields[1].AsInteger;
  end;
  Ferme(Q) ;

  If (Trouve = True) and (num <> No) then
  begin
 	  LastError := 1;
   	LastErrorMsg := 'Cette action est déjà définie dans le processus ';
    SetFocusControl('PRD_PROCESTYPE');
    exit;
 	end;

  // Recherche cohérence type processus et traitement
  If (Protype = 'ENT') and (Copy(Trt, 1, 1) <> 'E') then
  begin
 	  LastError := 1;
   	LastErrorMsg := 'Cette action est incompatible avec le type de processus ';
    SetFocusControl('PRD_PROCESACTION');
    exit;
 	end;
  If (Protype = 'SOR') and (Copy(Trt, 1, 1) <> 'S') then
  begin
 	  LastError := 1;
   	LastErrorMsg := 'Cette action est incompatible avec le type de processus ';
    SetFocusControl('PRD_PROCESACTION');
    exit;
 	end;

end ;

procedure TOM_PROCESSUSDET.OnLoadRecord ;
var
  Ordre : String;
  num : Integer;
  Q : TQuery;

begin
  Inherited ;
  SetField('PRD_PREDEFINI',predef);
  SetField('PRD_NODOSSIER',dossier);
  SetField('PRD_CODEPROC', code);

  AccesPredefini('TOUS', CEG, STD, DOS);
  LectureSeule := FALSE;
  if (Getfield('PRD_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
  end
  else
    if (Getfield('PRD_PREDEFINI') = 'STD') then
    begin
      LectureSeule := (STD = False);
      PaieLectureSeule(TFFiche(Ecran), (STD = False));
    end
    else
      if (Getfield('PRD_PREDEFINI') = 'DOS') then
      begin
        LectureSeule := (DOS = False);  //PT2
        PaieLectureSeule(TFFiche(Ecran), (DOS = False));    //PT2
        SetControlEnabled('PRD_PREDEFINI', False);
        SetControlEnabled('PRD_NODOSSIER', False);
        SetControlEnabled('PRD_CODEPROC', False);
        SetControlEnabled('PRD_NOORDRE', False);
        // SetControlVisible('BINSERT',True);
      end;

    // Appel fiche en insertion
    if DS.State in [dsInsert] then
    begin
      num := 1;
      LectureSeule := FALSE;
      PaieLectureSeule(TFFiche(Ecran), False);
      SetControlEnabled('PRD_PREDEFINI', False);
      SetControlEnabled('PRD_NODOSSIER', False);
      SetControlEnabled('PRD_CODEPROC', False);
	    Ordre := 'SELECT max (PRD_NOORDRE) FROM PROCESSUSDET WHERE PRD_CODEPROC="';
    	Ordre := Ordre + Code + '"';
  	  Q:= OpenSQL(Ordre,TRUE) ;
      if not q.eof then
        num := Q.Fields[0].AsInteger + 1;
  	  Ferme(Q) ;
      SetField('PRD_NOORDRE',num);
  end;

end ;

procedure TOM_PROCESSUSDET.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_PROCESSUSDET.OnArgument ( S: String ) ;
var
  BtnIns : TToolBarButton97;
  nobid,trt : String;
begin
  Inherited ;
  Arg := S;
  Trt := Trim(ReadTokenPipe(Arg,';')) ;
  predef:=Trim(ReadTokenPipe(Arg,';')) ;
  dossier:=Trim(ReadTokenPipe(Arg,';')) ;
  code := Trim(ReadTokenPipe(Arg,';')) ;
  nobid := Trim(ReadTokenPipe(Arg,';')) ;

  protype := Trim(ReadTokenPipe(Arg,';')) ;

  BtnIns := TToolBarButton97(GetControl('BINSERT'));
  if BtnIns <> nil then BtnIns.OnClick := BTnInsClick;
end ;

procedure TOM_PROCESSUSDET.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_PROCESSUSDET.OnCancelRecord ;
begin
  Inherited ;
end ;

//  Insertion d'une ligne
procedure TOM_PROCESSUSDET.BTnInsClick(Sender: TObject);
begin
  Inherited ;
end;

Initialization
  registerclasses ( [ TOM_PROCESSUSDET ] ) ;
end.
