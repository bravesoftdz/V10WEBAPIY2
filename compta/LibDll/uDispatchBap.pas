unit uDispatchBap;

interface

uses
  classes, UTOB, HEnt1;

procedure InitApplication;
function myDispatch(Action, Param : String ; RequestTOB : TOB) : TOB ;
function myProcCalcEdt(sf, sp : string) : variant;

implementation

uses
  {$IFDEF VER150} Variants,{$ENDIF} uFichierLog,
  sysutils, eSession, eDispatch, uLibBonAPayer, CPENVOYERMAIL, CPCREERMAIL, CPETAPEBAP,EntPGI,Ent1 ;

{---------------------------------------------------------------------------------------}
procedure InitApplication;
{---------------------------------------------------------------------------------------}
begin
  ProcDispatch := myDispatch;
  SetProcCalcEdt(myProcCalcEdt);
end;

{---------------------------------------------------------------------------------------}
function myDispatch(Action, Param : string; RequestTOB : TOB) : TOB;
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  {JP Pour le debugage : ligne de paramètre d'exécution à lancer une fois que les fichiers
        d'entrée et de sortie ont été "récupérés -i in.tmp -o out.tmp -a cloture -p aucun".
        C'est à quoi sert la boucle sans fin ci-dessous
    ;}
//  while True do begin end;
  // Init variable Halley
(* Modif GP : c'est compliqué : j'explique après la modif...
  InitAGL;
//  ChargeMagHalley ;
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
  Result := TOB.Create('', nil, -1) ;
  try
    ObjLog := TObjFichierLog.Create(UpperCase(Action));
    
    try
      try
        if UpperCase(Action) = psrv_Alerte then
          ch := TCreerMail.CreerMails(RequestTob)
        else if UpperCase(Action) = psrv_EnvoiMail then
          ch := TEnvoyerMail.EnvoyerMails(RequestTob)
        else if UpperCase(Action) = psrv_Etape then
          ch := TCreerEtapeBap.CreerEtape(RequestTob)
        else
          ch := TraduireMemoire('Pas d''action de défini');
      finally
        ObjLog.GenereFichier;
      end;
    finally
      if Assigned(ObjLog) then FreeAndNil(ObjLog);
    end;
  except
    on E : Exception do begin
      ch := TraduireMemoire('Erreur lors de l''exécution de la tâche ') + Action + #13#13 + E.Message;

    end;
  end;

  Result.AddChampSupValeur('ERROR', ch);
end;

{---------------------------------------------------------------------------------------}
function myProcCalcEdt(sf, sp : string) : variant;
{---------------------------------------------------------------------------------------}
begin
  Result := Unassigned;
end;


end.

