{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/06/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYYPLANNING_MUL ()
Mots clefs ... : TOF;YYYPLANNING_MUL
*****************************************************************}
Unit YYYPLANNING_MUL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HTB97,
     UTOF,
     AglInit,
     UTOB;

Type
  TOF_YYYPLANNING_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    Private
      BRecalculer:    TToolBarButton97;
      Procedure BRecalculerDoClick( Sender: TObject );

  end ;

Implementation
Uses
  YPlanning;     //Pour la fonction RecalculYPL;

procedure TOF_YYYPLANNING_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_YYYPLANNING_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_YYYPLANNING_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_YYYPLANNING_MUL.OnLoad ;
begin
  Inherited ;
end ;

Procedure TOF_YYYPLANNING_MUL.OnArgument( S : String );
Begin
  Inherited ;
  BRecalculer := TToolBarButton97( GetControl( 'BRECALCULER' ) );
  BRecalculer.OnClick := BRecalculerDoClick;
End ;

Procedure TOF_YYYPLANNING_MUL.BRecalculerDoClick ( Sender: TObject );
Var
  vTob      : Tob;
  vSt: String;

Begin
  If PGIAsk( 'Etes-vous sûr de vouloir recalculer le planning unifié ?', Ecran.Caption ) <> mrYes Then Exit;
  SourisSablier;

  //GA
  vSt := 'SELECT APL_ETATLIGNE, APL_RESSOURCE, APL_GUID, APL_HEUREDEB_PLA, APL_HEUREFIN_PLA, APL_LIBELLEPLA ';
  vSt := vST + ' FROM AFPLANNING WHERE APL_TYPEPLA = "PLA"';

  //20080506_C.B_GA12636
  // stopper le traitement
  if RecalculYPL(cStTableGa, vSt) then
  begin
    //GRC
    vSt := 'select rac_libelle, rac_etataction, rai_guid, rai_ressource,';
  //  vSt := vSt + ' iif(rac_heureaction="'+UsDateTime(iDate1900)+'" ,rac_dateaction+rac_heureaction,rac_dateaction+rac_heureaction+2) AS DATEDEBUT,';
  //  vSt := vSt + ' iif(rac_heureaction="'+UsDateTime(iDate1900)+'" ,rac_dateaction+rac_heureaction +cast((rac_dureeaction/59.9999)/24 as datetime), rac_dateaction+rac_heureaction+2 +cast((rac_dureeaction/59.9999)/24 as datetime)) AS DATEFIN';
    vSt := vSt + 'rac_dateaction,rac_heureaction,rac_dureeaction';
    vSt := vSt + ' FROM ACTIONINTERVENANT';
    vSt := vSt + ' left join actions on rai_auxiliaire = rac_auxiliaire and rai_numaction = rac_numaction';
    vSt := vSt + ' left join paractions on rac_typeaction = rpa_typeaction';
    vSt := vSt + ' WHERE rpa_planifiable = "X" and rac_etataction <> "ANU"';

    vTob := Tob.Create('#' + ' ACTIONINTERVENANT', nil, -1);
    try
      vTob.LoadDetailDBFromSql('#' + ' ACTIONINTERVENANT', vSt);
      if vTob.detail.count > 0 then
        RecalculYPL(cStTableGrc, '', vTob);
    finally
      vTob.Free;
    end;
  //  RecalculYPL(cStTableGrc, vSt);

    //BUREAU
    vSt := 'select JEV_GUIDEVT, JEV_DATE, JEV_DATEFIN, ';
    vSt := vSt + ' JEV_EVTLIBELLE, JEV_EVTLIBABREGE, JEV_USER1 ';
    vSt := vSt + ' FROM JUEVENEMENT ';
    vSt := vSt + ' WHERE JEV_FAMEVT="ACT" AND JEV_DOMAINEACT<>"JUR" AND JEV_OCCURENCEEVT<>"REC"';

    //20080506_C.B_GA12636
    // stopper le traitement
    if RecalculYPL(cStTableBureau, vSt) then
    begin
      //PAIE
      vSt:= 'SELECT PCN_GUID, PCN_DATEDEBUTABS, PCN_DATEFINABS, PCN_LIBELLE '+
            ' FROM ABSENCESALARIE WHERE'+
            ' PCN_ETATPOSTPAIE<>"NAN" AND'+
            ' (PCN_TYPEMVT="ABS" OR'+
            ' (PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI" AND PCN_MVTDUPLIQUE="-"))';
      RecalculYPL(cStTablePaie, vSt);

      SourisNormale;
      AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
    end;       
  end;
End;

procedure TOF_YYYPLANNING_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_YYYPLANNING_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_YYYPLANNING_MUL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_YYYPLANNING_MUL ] ) ; 
end.
