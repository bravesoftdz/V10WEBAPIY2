unit ZLibAuto;

interface

uses

 Classes,
 {$IFNDEF EAGLCLIENT}
 {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
 {$ENDIF}
 SysUtils,
 // composant AGL
 HCtrls,
 // Lib
 Ent1,
 Hent1,
 formule,   // pour GFormule
 ed_tools, // pour le VideListe
 // AGL
 LookUp,
// SaisUtil,
 UTOFLOOKUPTOB,
 stdctrls,
 UTOB ;

type

 TZLibAuto = class
  FTOB    : TOB;
  FTOBEcr : TOB;
 public
  constructor Create;
  destructor Destroy; override;

  procedure  Load;

  function RechercheEnBase( vEdit : TEdit ; vStLibelle , vStJournal , vStNature : string ; vTOBEcr : TOB ) : TOB;

 end; // class

implementation

{ TZLibAuto }

constructor TZLibAuto.Create;
begin
 FTOB := TOB.Create('',nil,-1);
end;

destructor TZLibAuto.Destroy;
begin
 if assigned(FTOB) then FTOB.Free;
 FTOB := nil;
 inherited;
end;

procedure TZLibAuto.Load;
var
 lQ : TQuery;
begin

 lQ := nil;

 try

  lQ := OpenSQL('select * from REFAUTO',true);
  FTOB.LoadDetailDB('REFAUTO','','',lQ,true);

 finally
  if assigned(lQ) then Ferme(lQ);
 end; // try

end;


function TZLibAuto.RechercheEnBase( vEdit : TEdit ; vStLibelle , vStJournal , vStNature : string ; vTOBEcr : TOB ) : TOB;
var
 lTOBLigne        : TOB;
 lTOBResult       : TOB;
 lTOBLigneResult  : TOB;
 lTOB             : TOB;
 i                : integer;
begin

 result     := TOB.Create('REFAUTO',nil,-1);
 lTOBResult := TOB.Create('',nil,-1);
 FTOBEcr    := vTOBEcr;

 try

 // recherche du libelle dans les libelle auto
 if vStLibelle <> '' then
 for i:= 0 to FTOB.Detail.Count - 1 do
  begin
   lTOBLigne := FTOB.Detail[i];
   if ( lTOBLigne.GetValue('RA_JOURNAL') = vStJournal )            and
      ( lTOBLigne.GetValue('RA_NATUREPIECE') = vStNature )         and
      ( pos(UpperCase(vStLibelle),UpperCase(lTOBLigne.GetValue('RA_FORMULELIB '))) = 1 ) then
    begin
     lTOBLigneResult := TOB.Create('REFAUTO',lTOBResult,-1);
     lTOBLigneResult.Dupliquer(lTOBLigne,false,true);
    end; // if
  end; // for

 // on n'a pas trouver de libelle auto -> on renvoie tous les libelles auto
 if lTOBResult.Detail.Count = 0 then
  for i:= 0 to FTOB.Detail.Count - 1 do
   begin
    lTOBLigne := FTOB.Detail[i];
    if ( lTOBLigne.GetValue('RA_JOURNAL') = vStJournal )     and
       ( lTOBLigne.GetValue('RA_NATUREPIECE') = vStNature ) then
     begin
      lTOBLigneResult := TOB.Create('REFAUTO',lTOBResult,-1);
      lTOBLigneResult.Dupliquer(lTOBLigne,false,true);
     end; // if
   end; // for

  if lTOBResult.Detail.Count = 1 then
   result.Dupliquer(lTOBResult.Detail[0],false,true)
    else
     begin

      if ( lTOBResult.Detail.Count > 0) then
       begin
        lTOB := LookUpTob ( vEdit ,
                            lTOBResult ,
                            'Libellé pour : ' + vStLibelle ,
// GP le 18/08/2008 N° 22611                            'RA_CODE;RA_FORMULELIB',
                            'RA_CODE;RA_LIBELLE',
                            'Code;Libellé' );

        if assigned(lTOB) then
         result.Dupliquer(lTOB,false,true)
          else
           begin
            result.free;
            result := nil;
           end;

       end; // if

     end; // if

 finally
  lTOBResult.Free;
 end;

end;

end.
