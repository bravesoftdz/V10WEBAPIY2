unit uDispatchCloture;

interface

uses
  classes,
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  UTOB,
  ULibCloture,
  ULibClotureAna,
  HEnt1,
  Ent1
  ;

procedure InitApplication;
function myDispatch(Action, Param : String ; RequestTOB : TOB) : TOB ;
function myProcCalcEdt(sf, sp : string) : variant;

implementation

uses
  sysutils,
//DBTables, HEnt1,HCtrls,
  eSession, eDispatch, EntPGI;

procedure InitApplication;
begin
  ProcDispatch := myDispatch;
  SetProcCalcEdt(myProcCalcEdt) ;
end;

function myDispatch(Action, Param : String ; RequestTOB : TOB) : TOB ;
var ClotureProcess    : TTraitementCloture ;
    ClotureAnaProcess : TTraitementClotureAna ;
begin

  {JP Pour le debugage : ligne de paramètre d'exécution à lancer une fois que les fichiers
        d'entrée et de sortie ont été "récupérés -i in.tmp -o out.tmp -a cloture -p aucun".
        C'est à quoi sert la boucle sans fin ci-dessous}
   //while True do begin end ;
  // Init variable Halley
(* Modif GP : c'est compliqué : j'explique après la modif...
  InitAGL;
    // Init variable Halley
  InitLaVariableHalley;
  ChargeMagHalley;
*)
  InitLaVariableEntPgi;
  VH ;
(* 27/06/2008
Explication :
Théoriquement, initAGL (HENT1) est sensé faire un appel à initProcAgl (procédure typée); Laquelle InitProcAGL est positionnée
dans l'initialization de ENTPGI (de Commun\lib) via :   InitProcAGL := InitialisationAGL ;

InitialisationAGL lance InitVariablesPGI qui fait entre autre :
InitLaVariableHalley;
InitLaVariableEntPgi;

InitLaVariableEntPgi fait :
  VH_EntPgi:=LaVariableEntPgi.Create ;
  VH_EntPgi.TobTableAlertes := tMemoryTob.Create('_YTABLEALERTES_','', LoadTableAlertes);
  VH_EntPgi.TobAlertes := tMemoryTob.Create('_YALERTES_','', LoadAlertes);
  VH_EntPgi.TobTablesLiees := tMemoryTob.Create('_YTABLEALERTELIEES_','', LoadTablesLiees);

ce qui permet à d'autres procédure de ENTPGI (comme ChargeParamsPGI) de ne pas planter sur des instructions du type :
  VH_EntPgi.TobTableAlertes.ClearDetail;

Ok tout va bien.

MAIS :
Si on ne fait plus appel à InitLaVariableEntPgi, alors un plantage se fera sur ChargeParamsPGI
à cause de VH_EntPgi.TobTableAlertes.ClearDetail (Pointeur à NIL)

Pour ne pas faire appel à  InitLaVariableEntPgi, il suffit que InitAGL ne fasse plus appel à initProcAGL, etc...

Et ce cas se produit depuis la version AGL 7.0.26.12 en compil IFDEF EAGLSERVEUR !

Donc, le processServeur plante.

La correction faite pour l'instant est provisoire et doit être revue.

*)

  if RequestTOB.GetValue('ctxPCL') = True then
  V_PGI.PGIContexte := V_PGI.PGIContexte + [ctxPCL];
  // -------------------
  // Fonction de Cloture
  // -------------------
  if UpperCase(Action) = 'CLOTURE' then
    begin
    // Instanciation processus
    ClotureProcess := TTraitementCloture.Create( nil, // pas d'écran
                                                 RequestTOB.GetValue('Exo1'),
                                                 RequestTOB.GetValue('Exo2'),
                                                 RequestTOB.GetValue('CloDef'),
                                                 RequestTOB.GetValue('Auto')
                                                ) ;
    // Paramétrage
    ClotureProcess.SetModeServeur( True ) ;
    // Traitement
    result := ClotureProcess.ExecuteCloture( RequestTOB ) ;   // result doit être virtuelle, mais elle peut avoir des filles réelles
    ClotureProcess.Free ;
    end
  // -----------------------------------
  // Fonction de d'Annulation de cloture
  // -----------------------------------
  else if UpperCase(Action) = 'DECLOTURE' then
    begin
    // Instanciation processus
    ClotureProcess := TTraitementCloture.Create( nil, // pas d'écran
                                                 RequestTOB.GetValue('Exo1'),
                                                 RequestTOB.GetValue('Exo2'),
                                                 RequestTOB.GetValue('CloDef'),
                                                 RequestTOB.GetValue('Auto')
                                                ) ;
    // Paramétrage
    ClotureProcess.SetModeServeur( True ) ;
    // Traitement
    result := ClotureProcess.ExecuteAnnuleCloture( RequestTOB ) ;   // result doit être virtuelle, mais elle peut avoir des filles réelles
    ClotureProcess.Free ;
    end
  // -------------------
  // Fonction de Cloture Analytique
  // -------------------
  else if UpperCase(Action) = 'CLOTUREANA' then
    begin
    // Instanciation processus
    ClotureAnaProcess := TTraitementClotureAna.Create( nil, // pas d'écran
                                                       True // Mode serveur
                                                      ) ;
    // Traitement
    result := ClotureAnaProcess.ExecuteClotureAna( RequestTOB ) ;   // result doit être virtuelle, mais elle peut avoir des filles réelles
    ClotureAnaProcess.Free ;
    end
  // -----------------------------------
  // Fonction de d'Annulation de cloture
  // -----------------------------------
  else if UpperCase(Action) = 'DECLOTUREANA' then
    begin
    // Instanciation processus
    // Instanciation processus
    ClotureAnaProcess := TTraitementClotureAna.Create( nil, // pas d'écran
                                                       True // Mode serveur
                                                      ) ;
    // Traitement
    result := ClotureAnaProcess.ExecuteAnnuleClotureAna( RequestTOB ) ;   // result doit être virtuelle, mais elle peut avoir des filles réelles
    ClotureAnaProcess.Free ;
    end
  else
    result := TOB.Create('', nil, -1) ;

 result.AddChampSupValeur('ERROR', '');// créer au moins (au plus ?) ce champ
  // ne pas abuser des ERROR <> '' : elles sont transmises à l'administrateur...
end;

function myProcCalcEdt(sf, sp : string) : variant;
begin
  // init, au cas où on n'a pas la fonction
  result := UnAssigned; // et pas autre chose
  // traitement ici :
  // ...
  // result := ''; si et seulement si le champ doit être blanc
end;


end.

