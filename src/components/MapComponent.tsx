import { useEffect, useState } from 'react';
import { MapContainer, TileLayer, Marker, Popup } from 'react-leaflet';
import 'leaflet/dist/leaflet.css';
import L from 'leaflet';
import { AdsResult } from '../App';

// @ts-ignore
import icon from 'leaflet/dist/images/marker-icon.png';
// @ts-ignore
import iconShadow from 'leaflet/dist/images/marker-shadow.png';

let DefaultIcon = L.icon({
  iconUrl: icon,
  shadowUrl: iconShadow,
  iconSize: [25, 41],
  iconAnchor: [12, 41]
});
L.Marker.prototype.options.icon = DefaultIcon;

interface MapComponentProps {
  results: AdsResult[];
}

export default function MapComponent({ results }: MapComponentProps) {
  const [geocodedResults, setGeocodedResults] = useState<(AdsResult & { lat: number; lng: number })[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let active = true;
    const topResults = results.slice(0, 100); // Limit to 100 to avoid rate limits
    
    const geocode = async () => {
      setLoading(true);
      const gResults = [];
      for (const res of topResults) {
        if (!active) break;
        if (!res.adresse || res.adresse.trim() === '') continue;
        try {
          const apiRes = await fetch(`https://api-adresse.data.gouv.fr/search/?q=${encodeURIComponent(res.adresse)}&limit=1`);
          const data = await apiRes.json();
          if (data.features && data.features.length > 0) {
            const [lng, lat] = data.features[0].geometry.coordinates;
            gResults.push({ ...res, lat, lng });
          }
        } catch (e) {
          console.error('err geocoding', e);
        }
      }
      if (active) {
        setGeocodedResults(gResults);
        setLoading(false);
      }
    };
    
    geocode();
    return () => { active = false; };
  }, [results]);

  if (loading) {
    return (
      <div className="h-[500px] w-full rounded-xl flex flex-col items-center justify-center bg-slate-100 text-slate-500 border border-slate-200">
        <div className="w-8 h-8 border-4 border-slate-300 border-t-blue-500 rounded-full animate-spin mb-4"></div>
        <p className="text-sm font-medium">Géocodage des adresses en cours (max 100 résultats)...</p>
      </div>
    );
  }

  if (geocodedResults.length === 0) {
    return (
      <div className="h-[500px] w-full rounded-xl flex items-center justify-center bg-slate-100 text-slate-500 border border-slate-200">
        <p className="text-sm font-medium">Aucune coordonnée trouvée pour ces résultats.</p>
      </div>
    );
  }

  const center = geocodedResults.length > 0 
    ? [geocodedResults[0].lat, geocodedResults[0].lng] 
    : [46.603354, 1.888334];

  return (
    <div className="h-[500px] w-full rounded-xl overflow-hidden border border-slate-200 shadow-sm relative z-0">
      <MapContainer center={center as any} zoom={12} scrollWheelZoom={true} style={{ height: '100%', width: '100%' }}>
        <TileLayer
          attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
          url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
        />
        {geocodedResults.map((res, i) => (
          <Marker key={`${res.id}-${i}`} position={[res.lat, res.lng]}>
            <Popup>
               <div className="font-sans min-w-[200px]">
                 <span className="text-xs font-bold text-slate-500 bg-slate-100 px-2 py-1 rounded inline-block mb-2">
                   {res.statut}
                 </span>
                 <h4 className="font-bold text-sm text-slate-800 mb-1 leading-tight">{res.type || 'Permis'}</h4>
                 <p className="text-xs text-blue-600 font-bold mb-2">{res.demandeur || 'Non renseigné'}</p>
                 <p className="text-xs text-slate-500 mb-2 leading-relaxed">{res.adresse}</p>
                 <div className="text-xs text-slate-700 bg-slate-50 p-2 rounded border border-slate-100">
                   {res.description}
                 </div>
               </div>
            </Popup>
          </Marker>
        ))}
      </MapContainer>
    </div>
  );
}
