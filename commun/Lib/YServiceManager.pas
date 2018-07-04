unit YServiceManager;

// Listes des interfaces à disposition

interface

uses
  ySouche_abs
{ GC_20090527_JTR_TD;14671_Début }
{$IFDEF GCGC}
  {$IFNDEF PGIMAJVER}
  , GCEcoContribution_abs
  {$ENDIF !PGIMAJVER}
{$ENDIF GCGC}
  ;
{ GC_200905127_JTR_TD;14671_Fin }

Type
  TYServiceManager = class(Tobject)
    class function SoucheManager : IySouche ;
{ GC_20090527_JTR_TD;14671_Début }
{$IFDEF GCGC}
  {$IFNDEF PGIMAJVER}
    class function EcoContribManager : IgcEcoContribution ;
  {$ENDIF !PGIMAJVER}
{$ENDIF GCGC}
{ GC_200905127_JTR_TD;14671_Fin }
  end;

implementation

uses
  ySouche
{ GC_20090527_JTR_TD;14671_Début }
{$IFDEF GCGC} // GM
  {$IFNDEF PGIMAJVER}
  , gcEcoContribution
  {$ENDIF !PGIMAJVER}
{$ENDIF GCGC} // GM
  ;
{ GC_200905127_JTR_TD;14671_Fin }

class function TYServiceManager.SoucheManager : IySouche;
begin
  result := yISouche;
end;

{ GC_20090527_JTR_TD;14671_Début }
{$IFDEF GCGC}
  {$IFNDEF PGIMAJVER}
class function TYServiceManager.EcoContribManager : IgcEcoContribution;
begin
  result := gcIEcoContribution;
end;
  {$ENDIF !PGIMAJVER}
{$ENDIF GCGC}
{ GC_200905127_JTR_TD;14671_Fin }

end.
