unit ZMAJEcriture;

interface

uses StdCtrls, Controls, Classes,
 {$IFNDEF EAGLCLIENT}
 db,dbTables,
 {$ENDIF}
 forms, sysutils,  ComCtrls,
 HCtrls, HEnt1, HMsgBox,
 // VCL
 Dialogs,
 Windows,
 // Lib
 SaisComm, // pour le TFRM
 ULibEcriture,
 UVar,
 Ent1,
 ULibExercice,
 SAISUTIL, // pour RDevise
 TZ, // pour le TZF
 UTOB;

type



 TZMAJEcriture = Class
 private
  FListComptes         : TList;
  FListJournaux        : TList;
  FBoChampsControle    : boolean;
  FStNomTable          : string;
  FStPrefixe           : string;
  FTTypeExo            : TTypeExo;

  function ControlePresenceChamps : boolean;
  function ControlePrefixe ( vTOBLigneEcr : TOB ) : boolean;
  function ControleStructureTOBEcr(vTOBEcr: TOB): boolean;
  function ControleStructureTOBLigneEcr(vTOBLigneEcr: TOB): boolean;
  function IndexOf ( vList : TList ; vStGeneral : string ) : integer;

 public

  constructor Create ( vStNomTable : string );
  destructor  Destroy; override;

  function    AddLigneEcr ( vTOBLigneEcr : TOB ) : boolean;
  function    AddPieceEcr ( vTOBEcr : TOB )      : boolean;
  function    AddBorEcr   ( vTOBEcr : TOB )      : boolean;

  function    Save : boolean;

 end; // class

implementation

{ TZTresoVersEcr }

constructor TZMAJEcriture.Create ( vStNomTable : string );
begin
  FListComptes         := TList.Create;
  FListJournaux        := TList.Create;
  FBoChampsControle    := false;
  FStNomTable          := vStNomTable;
  FStPrefixe           := TableToPrefixe( vStNomTable );
  FTTypeExo            := CGetTypeExo;
end;

destructor TZMAJEcriture.Destroy;
var
 i : integer;
begin

 for i := 0 to FListComptes.Count - 1 do
   dispose(FListComptes.Items[i]);

 for i := 0 to FListJournaux.Count - 1 do
   dispose(FListJournaux.Items[i]);

 FListComptes.free;
 TList.Create.Free;

 inherited;

end;

function TZMAJEcriture.ControlePresenceChamps : boolean;
begin
 result := true;
end;

function TZMAJEcriture.ControlePrefixe ( vTOBLigneEcr : TOB ) : boolean;
begin
 result := TableToPrefixe ( vTOBLigneEcr.NomTable ) = FStPrefixe;
end;

function TZMAJEcriture.ControleStructureTOBLigneEcr ( vTOBLigneEcr : TOB ) : boolean;
begin
 result := assigned( vTOBLigneEcr ) and ControlePrefixe ( vTOBLigneEcr ) ;
end;


function TZMAJEcriture.ControleStructureTOBEcr ( vTOBEcr : TOB ) : boolean;
begin

  result      := assigned( vTOBEcr )                   and
                 assigned( vTOBEcr.Detail)             and
                 ( vTOBEcr.Detail.Count > 0 )          and
                 ControlePrefixe ( vTOBEcr.Detail[0] ) ;

end;

function TZMAJEcriture.IndexOf ( vList : TList ; vStGeneral : string ) : integer;
var
 i       : integer;
 lpTFRM  : pTFRM;
begin

 result := -1;

 for i := 0 to vList.Count - 1 do
  begin
   lpTFRM :=  vList.Items[i];
   if lpTFRM^.Cpt = vStGeneral then
    begin
     result := i;
     break;
    end; // if
  end; // for

end;


function TZMAJEcriture.AddBorEcr( vTOBEcr : TOB) : boolean;
begin

end;

function TZMAJEcriture.AddLigneEcr( vTOBLigneEcr : TOB) : boolean;

 function Ajout ( vList : TList ; fb : TFichierBase ) : boolean;
 var
   lInIndex : integer;
   lpTFRM   : pTFRM;
  begin

   result   := true;

   if fb = fbGene then
    lInIndex := IndexOf ( vList , vTOBLigneEcr.GetValue( FStPrefixe + '_GENERAL' ) )
     else
      lInIndex := IndexOf ( vList , vTOBLigneEcr.GetValue( FStPrefixe + '_JOURNAL' ) );

   if lInIndex = - 1 then
    begin

     New(lpTFRM);

     FillChar ( lpTFRM^ , sizeOf(lpTFRM^) , #0 );

     if fb = fbGene then
      lpTFRM^.Cpt := vTOBLigneEcr.GetValue( FStPrefixe + '_GENERAL' )
       else
        lpTFRM^.Cpt := vTOBLigneEcr.GetValue( FStPrefixe + '_JOURNAL' );

     AttribParamsNewP ( lpTFRM ,
                       vTOBLigneEcr.GetValue( FStPrefixe + '_DEBIT' )  ,
                       vTOBLigneEcr.GetValue( FStPrefixe + '_CREDIT' ) ,
                       FTTypeExo
                      );
     vList.Add ( lpTFRM );

    end
     else
      begin

       lpTFRM :=  vList.Items[lInIndex];

       AttribParamsAjoutP ( lpTFRM ,
                           vTOBLigneEcr.GetValue( FStPrefixe + '_DEBIT' )  ,
                           vTOBLigneEcr.GetValue( FStPrefixe + '_CREDIT' ) ,
                           FTTypeExo
                          );
       if ( lpTFRM^.DateD < vTOBLigneEcr.GetValue( FStPrefixe + '_DATECOMPTABLE' ) ) then
        begin
         lpTFRM.Deb   := vTOBLigneEcr.GetValue( FStPrefixe + '_DEBIT' );           // debit du denier mouvement
         lpTFRM.Cre   := vTOBLigneEcr.GetValue( FStPrefixe + '_CREDIT' );          // credit du dernier mouvement
         lpTFRM.DateD := vTOBLigneEcr.GetValue( FStPrefixe + '_DATECOMPTABLE' );   // date comptable
         lpTFRM.NumD  := vTOBLigneEcr.GetValue( FStPrefixe + '_NUMEROPIECE' );
         lpTFRM.LigD  := vTOBLigneEcr.GetValue( FStPrefixe + '_NUMLIGNE' );
        end; // if

      end; // if

   result   := false;

  end; // procedure


begin

 result := Ajout ( FListComptes , fbGene );
 if result then
  result := Ajout ( FListJournaux , fbJal );

end;

function TZMAJEcriture.AddPieceEcr( vTOBEcr : TOB) : boolean;
var
 i : integer;
begin

 result := false;

 for i := 0 to ( vTOBEcr.Detail.Count - 1 ) do
  begin
   if not AddLigneEcr( vTOBEcr.Detail[i] ) then
    break;
  end; // for

  result := true;

end;


function TZMAJEcriture.Save : boolean;

 procedure SaveList( vList : TList ; fb : TFichierBase );
 var
  i : integer;
  lTFRM   : TFRM;
  lpTFRM  : pTFRM;

  begin
   for i := 0 to FListComptes.Count - 1 do
    begin

     lpTFRM      :=  FListComptes.Items[i];
     lTFRM.Cpt   := lpTFRM^.Cpt;           // numero du compte
     lTFRM.Deb   := lpTFRM^.Deb;           // debit du denier mouvement
     lTFRM.Cre   := lpTFRM^.Cre;           // credit du dernier mouvement
     lTFRM.DateD := lpTFRM^.DateD;         // date du dernier mouvement
     lTFRM.DE    := lpTFRM^.DE;
     lTFRM.CE    := lpTFRM^.CE;
     lTFRM.DS    := lpTFRM^.DS;
     lTFRM.CS    := lpTFRM^.CS;
     lTFRM.DP    := lpTFRM^.DP;
     lTFRM.CP    := lpTFRM^.CP;

     if ExecReqMAJNew ( fb , false, false , lTFRM ) <> 1 then
      break;

    end; // for

  end; // procedure


begin

 result := false;

 SaveList( FListComptes , fbGene);
 SaveList( FListJournaux , fbJal);

 result := true;

end;

end.
