{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 06/07/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : INFOSPLANNING ()
Mots clefs ... : TOF;INFOSPLANNING
*****************************************************************}
Unit UtofInfosPlanning ;

Interface

Uses Classes,
     uTob,
     UTOF,
     EntRT
      ;

Type
  TOF_INFOSPLANNING = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    TobModelePlanning: Tob;
    function LoadTobParam(T:TOB): integer;
  end ;

Implementation

procedure TOF_INFOSPLANNING.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_INFOSPLANNING.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_INFOSPLANNING.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_INFOSPLANNING.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_INFOSPLANNING.OnArgument (S : String ) ;
var vTOBModelePLanningXML : TOB;
    vListe                : TStringList;
    vStream               : TStream;
begin
  Inherited ;
  TobModelePlanning   := TOB.Create('Les_modeles',Nil,-1) ;
  vListe := TStringList.Create;

  VH_RT.TobParamPlanning.Load;

  vTOBModelePLanningXML:=VH_RT.TobParamPlanning.FindFirst(['RPP_CODEPARAMPLAN'],['PL2'],TRUE) ;
  if vTOBModelePLanningXML <> Nil then
    begin

    // chargement de la liste
    vListe.Text := vTOBModelePLanningXML.GetValue('RPP_PARAMS');

    // transfert dans une stream
    vStream := TStringStream.Create(vListe.Text);

    // recuperation dans une tob virtuelle fTOBModelePlanning
    TOBLoadFromXMLStream(vStream, LoadTobParam);
    SetControlText ('TYPEACTIONS',TOBModelePlanning.detail[0].Getvalue('TYPEACTIONS'));
    /////////////////////////////////////:
    vStream.Free;
    end;
  vListe.Free;
  TobModelePlanning.free;
end ;

procedure TOF_INFOSPLANNING.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_INFOSPLANNING.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_INFOSPLANNING.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_INFOSPLANNING.LoadTobParam (T:TOB): integer;
var
  NewTob : Tob;
begin

  NewTOB := TOB.Create('fille_param', TOBModelePlanning, -1);
  try
    NewTob.Dupliquer(T.detail[0],True,True);
    result := 0;
  finally
    T.Free;
  end;
end;

Initialization
  registerclasses ( [ TOF_INFOSPLANNING ] ) ; 
end.
 
