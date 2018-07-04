unit ImOutiCB;

interface

type TImContrat = class
     private
     public
       ListeEcheances : TList;
       dtPremEch : TDateTime;
       dtDernEch : TDateTime;
       dtDebut : TDateTime;
       dtFin : TDateTime;
       procedure CalculEcheances;
     end;

implementation

procedure TImContrat.CalculEcheances;
begin
end;

end.
