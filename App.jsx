import React, { useState, useEffect, useRef } from 'react';
import { initializeApp } from 'firebase/app';
import { getAuth, signInAnonymously, signInWithCustomToken, onAuthStateChanged } from 'firebase/auth';
import { getFirestore, collection, onSnapshot, doc, addDoc, updateDoc, writeBatch } from 'firebase/firestore';
import { 
  Search, Plus, MapPin, History, Camera, X, Check, Barcode, 
  ArrowRightLeft, Package, AlertCircle, UploadCloud, FileSpreadsheet, 
  Image as ImageIcon, LayoutDashboard, PlusCircle
} from 'lucide-react';

// --- Inicialização Global ---
const appId = typeof __app_id !== 'undefined' ? __app_id : 'default-app-id';
const firebaseConfig = typeof __firebase_config !== 'undefined' ? JSON.parse(__firebase_config) : null;

let app, auth, db;
if (firebaseConfig) {
  app = initializeApp(firebaseConfig);
  auth = getAuth(app);
  db = getFirestore(app);
}

// Carrega dinamicamente a biblioteca de QR Code
const loadScannerLibrary = () => {
  return new Promise((resolve, reject) => {
    if (window.Html5QrcodeScanner) {
      resolve();
      return;
    }
    const script = document.createElement('script');
    script.src = 'https://unpkg.com/html5-qrcode';
    script.async = true;
    script.onload = resolve;
    script.onerror = reject;
    document.body.appendChild(script);
  });
};

// Carrega dinamicamente a biblioteca SheetJS (Excel)
const loadXlsxLibrary = () => {
  return new Promise((resolve, reject) => {
    if (window.XLSX) {
      resolve();
      return;
    }
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    script.async = true;
    script.onload = resolve;
    script.onerror = reject;
    document.body.appendChild(script);
  });
};

// --- Componentes Úteis ---

const Toast = ({ message, type, onClose }) => {
  useEffect(() => {
    const timer = setTimeout(onClose, 3000);
    return () => clearTimeout(timer);
  }, [onClose]);

  const bgColor = type === 'error' ? 'bg-red-500' : 'bg-emerald-600';
  
  return (
    <div className={`fixed top-4 left-1/2 -translate-x-1/2 w-11/12 max-w-md ${bgColor} text-white px-4 py-3 rounded-xl shadow-2xl flex items-center gap-3 z-50 animate-fade-in-down`}>
      {type === 'error' ? <AlertCircle size={22} className="flex-shrink-0" /> : <Check size={22} className="flex-shrink-0" />}
      <span className="font-medium text-sm sm:text-base leading-tight">{message}</span>
    </div>
  );
};

const BarcodeScannerModal = ({ onScan, onClose }) => {
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let scanner;
    const init = async () => {
      try {
        await loadScannerLibrary();
        setLoading(false);
        scanner = new window.Html5QrcodeScanner(
          "qr-reader",
          { fps: 10, qrbox: { width: 250, height: 100 }, aspectRatio: 1.0 },
          false
        );
        scanner.render((decodedText) => {
          if (scanner) scanner.clear();
          onScan(decodedText);
        }, () => {});
      } catch (err) {
        console.error("Falha ao carregar scanner", err);
      }
    };
    init();
    return () => { if (scanner) scanner.clear().catch(e => console.error(e)); };
  }, [onScan]);

  return (
    <div className="fixed inset-0 bg-black/90 z-50 flex flex-col items-center justify-center p-4 backdrop-blur-sm">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md overflow-hidden flex flex-col">
        <div className="p-4 border-b flex justify-between items-center bg-gray-50">
          <h3 className="font-bold text-lg flex items-center gap-2 text-gray-800">
            <Camera size={20} className="text-blue-600" /> Escanear Código
          </h3>
          <button onClick={onClose} className="p-2 bg-gray-200 hover:bg-gray-300 rounded-full transition-colors">
            <X size={20} className="text-gray-700" />
          </button>
        </div>
        <div className="p-2 sm:p-4 bg-black">
          {loading ? (
            <div className="h-64 flex items-center justify-center text-white">Carregando câmera...</div>
          ) : (
            <div id="qr-reader" className="w-full bg-black rounded-lg overflow-hidden border-2 border-gray-800"></div>
          )}
        </div>
        <div className="p-4 bg-gray-50 text-center text-sm text-gray-500">
          Aponte a câmera para o código de barras ou QR Code.
        </div>
      </div>
    </div>
  );
};

const HistoryModal = ({ asset, onClose }) => {
  if (!asset) return null;
  return (
    <div className="fixed inset-0 bg-black/60 z-40 flex items-end sm:items-center justify-center sm:p-4 backdrop-blur-sm">
      <div className="bg-white rounded-t-2xl sm:rounded-2xl shadow-2xl w-full max-w-lg h-[80vh] sm:h-auto sm:max-h-[80vh] flex flex-col animate-slide-up sm:animate-none">
        <div className="p-4 border-b flex justify-between items-center sticky top-0 bg-white rounded-t-2xl">
          <h3 className="font-bold text-lg flex items-center gap-2 text-gray-800 truncate pr-4">
            <History size={20} className="text-blue-600 flex-shrink-0" /> 
            <span className="truncate">{asset.name}</span>
          </h3>
          <button onClick={onClose} className="p-2 bg-gray-100 hover:bg-gray-200 rounded-full flex-shrink-0">
            <X size={20} />
          </button>
        </div>
        <div className="p-4 overflow-y-auto flex-1">
          <div className="mb-6 p-4 bg-blue-50 rounded-xl text-sm text-blue-900 border border-blue-100 flex items-start gap-3">
            <MapPin size={18} className="mt-0.5 text-blue-500 flex-shrink-0" />
            <div>
              <p className="font-semibold mb-1">Local Original de Cadastro</p>
              <p>{asset.initialLocation}</p>
            </div>
          </div>
          {(!asset.history || asset.history.length === 0) ? (
            <div className="text-center py-8 text-gray-400 flex flex-col items-center">
              <Package size={40} className="mb-3 opacity-20" />
              <p>Nenhuma transferência registrada.</p>
            </div>
          ) : (
            <div className="relative border-l-2 border-gray-200 ml-4 space-y-8 pb-4">
              {asset.history.slice().reverse().map((record, idx) => (
                <div key={idx} className="pl-6 relative">
                  <div className="absolute w-4 h-4 bg-blue-500 rounded-full -left-[9px] top-1 border-2 border-white shadow-sm"></div>
                  <div className="text-xs font-bold text-blue-600 uppercase tracking-wider mb-1">
                    {new Date(record.date).toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' })}
                  </div>
                  <div className="bg-gray-50 p-3 rounded-lg border border-gray-100 text-sm text-gray-700 shadow-sm mt-2">
                    Movido de <span className="font-bold text-gray-900">{record.from}</span> <br/>
                    para <span className="font-bold text-gray-900">{record.to}</span>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

// --- App Principal ---

export default function App() {
  const [user, setUser] = useState(null);
  const [assets, setAssets] = useState([]);
  const [loading, setLoading] = useState(true);
  
  // UI State
  const [searchQuery, setSearchQuery] = useState('');
  const [activeTab, setActiveTab] = useState('dashboard');
  const [toast, setToast] = useState(null);
  const [showScannerFor, setShowScannerFor] = useState(null);
  const [historyAsset, setHistoryAsset] = useState(null);

  // Pagination & Import
  const [importProgress, setImportProgress] = useState(null);
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 20; // Menor para mobile scrollar melhor

  // Forms
  const [addForm, setAddForm] = useState({ name: '', barcode: '', location: '', description: '', imageBase64: '' });
  const [transferForm, setTransferForm] = useState({ barcode: '', newLocation: '' });

  // 1. Auth & Firebase Init
  useEffect(() => {
    if (!auth) return;
    const initAuth = async () => {
      try {
        if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
          await signInWithCustomToken(auth, __initial_auth_token);
        } else {
          await signInAnonymously(auth);
        }
      } catch (error) {
        showToast("Erro de autenticação", "error");
      }
    };
    initAuth();
    const unsubscribe = onAuthStateChanged(auth, setUser);
    return () => unsubscribe();
  }, []);

  // 2. Fetch Assets
  useEffect(() => {
    if (!user || !db) return;
    const assetsRef = collection(db, 'artifacts', appId, 'public', 'data', 'assets');
    const unsubscribe = onSnapshot(assetsRef, (snapshot) => {
      const loadedAssets = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      // Ordenar do mais novo para o mais velho (opcional, melhor UX)
      loadedAssets.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
      setAssets(loadedAssets);
      setLoading(false);
    }, (error) => {
      showToast("Erro ao carregar dados", "error");
      setLoading(false);
    });
    return () => unsubscribe();
  }, [user]);

  useEffect(() => { setCurrentPage(1); }, [searchQuery]);

  const showToast = (message, type = 'success') => setToast({ message, type });

  // --- Função de Compressão de Imagem ---
  const compressImage = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        const img = new Image();
        img.onload = () => {
          const canvas = document.createElement('canvas');
          const MAX_WIDTH = 500; // Largura máxima otimizada
          const scaleSize = MAX_WIDTH / img.width;
          
          // Se a imagem for menor que 500px, mantém o tamanho original
          let width = img.width;
          let height = img.height;
          
          if (width > MAX_WIDTH) {
            width = MAX_WIDTH;
            height = img.height * scaleSize;
          }

          canvas.width = width;
          canvas.height = height;
          
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
          
          // Exporta como JPEG com qualidade 0.7 (70%)
          const compressedBase64 = canvas.toDataURL('image/jpeg', 0.7);
          resolve(compressedBase64);
        };
        img.onerror = (err) => reject(err);
        img.src = event.target.result;
      };
      reader.onerror = (err) => reject(err);
      reader.readAsDataURL(file);
    });
  };

  const handleImageCapture = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.type.startsWith('image/')) {
      showToast("Por favor, selecione uma imagem.", "error");
      return;
    }

    try {
      showToast("Processando imagem...", "success");
      const compressedData = await compressImage(file);
      setAddForm(prev => ({ ...prev, imageBase64: compressedData }));
      showToast("Imagem adicionada!");
    } catch (error) {
      console.error("Erro ao comprimir:", error);
      showToast("Erro ao processar imagem.", "error");
    }
  };

  // --- Handlers ---

  const handleAddAsset = async (e) => {
    e.preventDefault();
    if (!addForm.name || !addForm.barcode || !addForm.location) {
      showToast("Preencha os campos obrigatórios (*)", "error");
      return;
    }

    const exists = assets.some(a => a.barcode === addForm.barcode);
    if (exists) {
      showToast("Este código de barras já existe!", "error");
      return;
    }

    try {
      const assetsRef = collection(db, 'artifacts', appId, 'public', 'data', 'assets');
      await addDoc(assetsRef, {
        name: addForm.name,
        barcode: addForm.barcode,
        description: addForm.description,
        initialLocation: addForm.location,
        currentLocation: addForm.location,
        image: addForm.imageBase64 || null,
        createdAt: new Date().toISOString(),
        createdBy: user.uid,
        history: [] 
      });
      
      showToast("Ativo cadastrado com sucesso!");
      setAddForm({ name: '', barcode: '', location: '', description: '', imageBase64: '' });
      setActiveTab('dashboard');
    } catch (error) {
      showToast("Erro ao salvar ativo", "error");
    }
  };

  const handleTransferAsset = async (e) => {
    e.preventDefault();
    if (!transferForm.barcode || !transferForm.newLocation) {
      showToast("Preencha código e destino", "error");
      return;
    }
    const assetToTransfer = assets.find(a => a.barcode === transferForm.barcode);
    if (!assetToTransfer) return showToast("Ativo não encontrado", "error");
    if (assetToTransfer.currentLocation === transferForm.newLocation) return showToast("Ativo já está neste local", "error");

    try {
      const assetRef = doc(db, 'artifacts', appId, 'public', 'data', 'assets', assetToTransfer.id);
      const newHistoryEntry = {
        date: new Date().toISOString(),
        from: assetToTransfer.currentLocation,
        to: transferForm.newLocation,
        userId: user.uid
      };
      await updateDoc(assetRef, {
        currentLocation: transferForm.newLocation,
        history: [...(assetToTransfer.history || []), newHistoryEntry]
      });
      showToast("Transferência confirmada!");
      setTransferForm({ barcode: '', newLocation: '' });
      setActiveTab('dashboard');
    } catch (error) {
      showToast("Erro ao transferir", "error");
    }
  };

  const handleScan = (decodedText) => {
    if (showScannerFor === 'add') setAddForm(prev => ({ ...prev, barcode: decodedText }));
    else if (showScannerFor === 'transfer') setTransferForm(prev => ({ ...prev, barcode: decodedText }));
    else if (showScannerFor === 'search') setSearchQuery(decodedText);
    setShowScannerFor(null);
    showToast("Código lido com sucesso!");
  };

  const handleExcelUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true);
    setImportProgress({ current: 0, total: 0, status: 'Lendo arquivo Excel...' });
    
    try {
      await loadXlsxLibrary();
      const reader = new FileReader();
      reader.onload = async (event) => {
        try {
          setImportProgress({ current: 0, total: 0, status: 'Processando dados...' });
          const data = new Uint8Array(event.target.result);
          const workbook = window.XLSX.read(data, { type: 'array' });
          const worksheet = workbook.Sheets[workbook.SheetNames[0]];
          const rows = window.XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          
          const existingBarcodes = new Set(assets.map(a => a.barcode));
          const validNewItems = [];
          let duplicateCount = 0;

          for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (row && row.length >= 3) {
              const barcode = String(row[0] || '').trim();
              const name = String(row[1] || '').trim();
              const location = String(row[2] || '').trim();
              const description = String(row[3] || '').trim();

              if (!barcode || !name || !location) continue;
              
              if (existingBarcodes.has(barcode)) {
                duplicateCount++;
              } else {
                validNewItems.push({
                  name, barcode, description, image: null,
                  initialLocation: location, currentLocation: location,
                  createdAt: new Date().toISOString(), createdBy: user.uid, history: []
                });
                existingBarcodes.add(barcode);
              }
            }
          }

          if (validNewItems.length === 0) {
            showToast(`Nenhum item novo. ${duplicateCount} duplicatas ignoradas.`);
            setLoading(false); setImportProgress(null); e.target.value = null; return;
          }

          const CHUNK_SIZE = 400;
          let processedCount = 0;
          const assetsCollectionRef = collection(db, 'artifacts', appId, 'public', 'data', 'assets');

          for (let i = 0; i < validNewItems.length; i += CHUNK_SIZE) {
            const chunk = validNewItems.slice(i, i + CHUNK_SIZE);
            const batch = writeBatch(db);
            chunk.forEach(item => { batch.set(doc(assetsCollectionRef), item); });
            await batch.commit();
            processedCount += chunk.length;
            setImportProgress({ current: processedCount, total: validNewItems.length, status: `Salvando ${processedCount} de ${validNewItems.length}...` });
          }

          showToast(`${processedCount} itens salvos. ${duplicateCount} duplicados ignorados.`);
          setActiveTab('dashboard');
        } catch (err) {
          showToast("Erro ao processar planilha.", "error");
        } finally {
          setLoading(false); setImportProgress(null); e.target.value = null;
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (err) {
      showToast("Falha ao carregar suporte a Excel.", "error");
      setLoading(false); setImportProgress(null);
    }
  };

  // --- Filtros e Paginação ---
  const filteredAssets = assets.filter(a => {
    const q = searchQuery.toLowerCase();
    return (
      a.name?.toLowerCase().includes(q) || 
      a.barcode?.toLowerCase().includes(q) || 
      a.currentLocation?.toLowerCase().includes(q)
    );
  });
  const totalPages = Math.max(1, Math.ceil(filteredAssets.length / itemsPerPage));
  const currentAssets = filteredAssets.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage);

  // --- Renderização ---

  if (!firebaseConfig) return <div className="p-8 text-center text-red-500 font-bold">Configuração do Firebase ausente.</div>;

  return (
    <div className="min-h-screen bg-gray-50 text-gray-800 font-sans pb-24">
      {/* Header Fixo Mobile */}
      <header className="bg-blue-600 text-white shadow-md sticky top-0 z-30">
        <div className="max-w-md mx-auto px-4 py-4 flex justify-between items-center">
          <div className="flex items-center gap-2 text-xl font-bold tracking-wide">
            <Package size={26} />
            <span>Controle Ativos</span>
          </div>
          {activeTab === 'dashboard' && (
            <div className="text-sm font-semibold bg-white/20 px-3 py-1 rounded-full">
              {assets.length}
            </div>
          )}
        </div>
      </header>

      {/* Conteúdo Principal */}
      <main className="max-w-md mx-auto w-full p-4">
        
        {loading && (
          <div className="flex flex-col items-center justify-center py-32 space-y-4">
            <div className="animate-spin rounded-full h-12 w-12 border-4 border-blue-600 border-t-transparent"></div>
            <p className="text-gray-500 font-medium animate-pulse">Carregando base de dados...</p>
          </div>
        )}

        {/* --- DASHBOARD --- */}
        {!loading && activeTab === 'dashboard' && (
          <div className="space-y-4 animate-fade-in">
            {/* Buscador Mobile */}
            <div className="relative shadow-sm rounded-xl overflow-hidden bg-white border border-gray-200">
              <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-gray-400" size={20} />
              <input 
                type="text" 
                placeholder="Buscar por nome, código ou local" 
                className="w-full pl-12 pr-14 py-4 bg-transparent focus:outline-none text-base"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
              />
              <button 
                onClick={() => setShowScannerFor('search')}
                className="absolute right-2 top-1/2 -translate-y-1/2 p-2 bg-blue-50 text-blue-600 rounded-lg active:bg-blue-100 transition-colors"
              >
                <Camera size={20} />
              </button>
            </div>

            {/* Lista de Ativos Mobile-First */}
            <div className="flex flex-col gap-3 mt-4">
              {currentAssets.length === 0 ? (
                <div className="text-center py-16 text-gray-500 bg-white rounded-2xl border border-dashed border-gray-300">
                  <Package size={48} className="mx-auto mb-4 text-gray-300" />
                  <p className="font-medium text-lg">Nenhum ativo encontrado.</p>
                  <p className="text-sm mt-1 text-gray-400">Tente buscar por outro termo.</p>
                </div>
              ) : (
                currentAssets.map(asset => (
                  <div key={asset.id} className="bg-white rounded-2xl p-4 shadow-sm border border-gray-100 flex flex-col gap-3 relative">
                    <div className="flex gap-4">
                      {/* Foto Thumbnail */}
                      <div className="w-20 h-20 bg-gray-100 rounded-xl overflow-hidden flex-shrink-0 border border-gray-200 flex items-center justify-center">
                        {asset.image ? (
                          <img src={asset.image} alt={asset.name} className="w-full h-full object-cover" />
                        ) : (
                          <ImageIcon size={32} className="text-gray-300" />
                        )}
                      </div>
                      
                      {/* Dados principais */}
                      <div className="flex-1 min-w-0">
                        <h3 className="font-bold text-gray-900 text-lg leading-tight truncate mb-1">
                          {asset.name}
                        </h3>
                        <div className="flex items-center gap-1.5 text-gray-500 font-mono text-xs bg-gray-100 w-fit px-2 py-1 rounded">
                          <Barcode size={12} /> {asset.barcode}
                        </div>
                        <div className="flex items-start gap-1.5 mt-2 text-sm text-gray-700">
                          <MapPin size={16} className="text-red-500 flex-shrink-0 mt-0.5" />
                          <span className="leading-tight line-clamp-2">{asset.currentLocation}</span>
                        </div>
                      </div>
                    </div>

                    {/* Botão de Histórico */}
                    <div className="border-t border-gray-100 pt-3 mt-1 flex justify-between items-center">
                      <span className="text-xs text-gray-400 truncate w-32">
                        {new Date(asset.createdAt).toLocaleDateString('pt-BR')}
                      </span>
                      <button 
                        onClick={() => setHistoryAsset(asset)}
                        className="flex items-center gap-1 text-blue-600 text-sm font-semibold bg-blue-50 px-3 py-1.5 rounded-lg active:bg-blue-100"
                      >
                        <History size={16} /> Histórico
                      </button>
                    </div>
                  </div>
                ))
              )}
            </div>

            {/* Paginação Mobile */}
            {totalPages > 1 && (
              <div className="flex justify-between items-center bg-white p-2 rounded-2xl shadow-sm border border-gray-100 mt-6">
                <button 
                  onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                  disabled={currentPage === 1}
                  className="px-4 py-3 bg-gray-50 text-gray-700 rounded-xl active:bg-gray-200 disabled:opacity-30 font-medium w-24 text-center"
                >
                  Anterior
                </button>
                <div className="text-sm font-bold text-gray-800">
                  {currentPage} / {totalPages}
                </div>
                <button 
                  onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                  disabled={currentPage === totalPages}
                  className="px-4 py-3 bg-blue-50 text-blue-700 rounded-xl active:bg-blue-200 disabled:opacity-30 font-medium w-24 text-center"
                >
                  Próxima
                </button>
              </div>
            )}
          </div>
        )}

        {/* --- ADD ASSET --- */}
        {!loading && activeTab === 'add' && (
          <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden animate-slide-in">
            <div className="bg-blue-600/10 px-5 py-4 border-b border-blue-100">
              <h2 className="font-bold text-xl text-blue-900 flex items-center gap-2">
                <PlusCircle size={22} className="text-blue-600" /> Cadastrar Ativo
              </h2>
            </div>
            
            <form onSubmit={handleAddAsset} className="p-5 space-y-5">
              
              {/* Câmera / Foto */}
              <div className="flex flex-col items-center gap-3 pb-4 border-b border-gray-100">
                <div className="w-32 h-32 rounded-2xl bg-gray-100 border-2 border-dashed border-gray-300 overflow-hidden flex items-center justify-center relative shadow-inner">
                  {addForm.imageBase64 ? (
                    <>
                      <img src={addForm.imageBase64} alt="Preview" className="w-full h-full object-cover" />
                      <button 
                        type="button" 
                        onClick={() => setAddForm({...addForm, imageBase64: ''})}
                        className="absolute top-1 right-1 bg-red-500 text-white p-1 rounded-full shadow-md"
                      >
                        <X size={16} />
                      </button>
                    </>
                  ) : (
                    <ImageIcon size={40} className="text-gray-300" />
                  )}
                </div>
                <label className="flex items-center gap-2 bg-blue-50 text-blue-700 px-5 py-2.5 rounded-full font-semibold cursor-pointer active:bg-blue-100 transition-colors">
                  <Camera size={18} />
                  <span>{addForm.imageBase64 ? 'Trocar Foto' : 'Tirar Foto'}</span>
                  <input 
                    type="file" 
                    accept="image/*" 
                    capture="environment" // Força abrir a câmera traseira no mobile se disponível
                    className="hidden" 
                    onChange={handleImageCapture} 
                  />
                </label>
              </div>

              <div>
                <label className="block text-sm font-bold text-gray-700 mb-1.5">Código de Barras *</label>
                <div className="flex gap-2">
                  <input 
                    type="text" required
                    className="flex-1 px-4 py-3 bg-gray-50 border border-gray-200 rounded-xl focus:ring-2 focus:ring-blue-500 focus:bg-white outline-none font-mono text-lg"
                    value={addForm.barcode}
                    onChange={(e) => setAddForm({...addForm, barcode: e.target.value})}
                    placeholder="PAT-00123"
                  />
                  <button 
                    type="button"
                    onClick={() => setShowScannerFor('add')}
                    className="bg-gray-800 text-white p-3 rounded-xl flex items-center justify-center active:bg-gray-700"
                  >
                    <Barcode size={24} />
                  </button>
                </div>
              </div>

              <div>
                <label className="block text-sm font-bold text-gray-700 mb-1.5">Nome do Ativo *</label>
                <input 
                  type="text" required
                  className="w-full px-4 py-3 bg-gray-50 border border-gray-200 rounded-xl focus:ring-2 focus:ring-blue-500 focus:bg-white outline-none text-base"
                  value={addForm.name}
                  onChange={(e) => setAddForm({...addForm, name: e.target.value})}
                  placeholder="Ex: Notebook Dell Latitude"
                />
              </div>

              <div>
                <label className="block text-sm font-bold text-gray-700 mb-1.5">Localização Inicial *</label>
                <input 
                  type="text" required
                  className="w-full px-4 py-3 bg-gray-50 border border-gray-200 rounded-xl focus:ring-2 focus:ring-blue-500 focus:bg-white outline-none text-base"
                  value={addForm.location}
                  onChange={(e) => setAddForm({...addForm, location: e.target.value})}
                  placeholder="Ex: Sala de TI - Andar 2"
                />
              </div>

              <div>
                <label className="block text-sm font-bold text-gray-700 mb-1.5">Descrição (Opcional)</label>
                <textarea 
                  className="w-full px-4 py-3 bg-gray-50 border border-gray-200 rounded-xl focus:ring-2 focus:ring-blue-500 focus:bg-white outline-none text-base resize-none"
                  rows="3"
                  value={addForm.description}
                  onChange={(e) => setAddForm({...addForm, description: e.target.value})}
                  placeholder="Nº de série, estado de conservação, etc."
                ></textarea>
              </div>

              <div className="pt-2">
                <button 
                  type="submit"
                  className="w-full bg-blue-600 text-white px-6 py-4 rounded-xl font-bold shadow-lg shadow-blue-200 flex items-center justify-center gap-2 active:bg-blue-700 active:scale-[0.98] transition-all text-lg"
                >
                  <Check size={24} /> Salvar Ativo
                </button>
              </div>
            </form>
          </div>
        )}

        {/* --- TRANSFER ASSET --- */}
        {!loading && activeTab === 'transfer' && (
          <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden animate-slide-in">
            <div className="bg-amber-500/10 px-5 py-4 border-b border-amber-100">
              <h2 className="font-bold text-xl text-amber-900 flex items-center gap-2">
                <ArrowRightLeft size={22} className="text-amber-600" /> Transferência
              </h2>
            </div>
            
            <form onSubmit={handleTransferAsset} className="p-5 space-y-6">
              <div>
                <label className="block text-sm font-bold text-gray-700 mb-1.5">Código do Ativo *</label>
                <div className="flex gap-2">
                  <input 
                    type="text" required
                    className="flex-1 px-4 py-3 bg-gray-50 border border-gray-200 rounded-xl focus:ring-2 focus:ring-amber-500 focus:bg-white outline-none font-mono text-lg"
                    value={transferForm.barcode}
                    onChange={(e) => setTransferForm({...transferForm, barcode: e.target.value})}
                    placeholder="Escanear ou digitar"
                  />
                  <button 
                    type="button"
                    onClick={() => setShowScannerFor('transfer')}
                    className="bg-amber-100 text-amber-700 p-3 rounded-xl flex items-center justify-center active:bg-amber-200 border border-amber-200"
                  >
                    <Barcode size={24} />
                  </button>
                </div>
                
                {/* Preview Mobile */}
                {transferForm.barcode && assets.find(a => a.barcode === transferForm.barcode) && (
                  <div className="mt-3 p-3 bg-green-50 border border-green-200 rounded-xl text-sm flex gap-3 items-center">
                    <Check size={20} className="text-green-600 flex-shrink-0" />
                    <div>
                      <strong className="text-green-900">{assets.find(a => a.barcode === transferForm.barcode).name}</strong> <br/>
                      <span className="text-green-700">Atual: {assets.find(a => a.barcode === transferForm.barcode).currentLocation}</span>
                    </div>
                  </div>
                )}
              </div>

              <div>
                <label className="block text-sm font-bold text-gray-700 mb-1.5">Novo Local (Destino) *</label>
                <div className="relative">
                  <MapPin className="absolute left-4 top-1/2 -translate-y-1/2 text-gray-400" size={20} />
                  <input 
                    type="text" required
                    className="w-full pl-12 pr-4 py-3 bg-gray-50 border border-gray-200 rounded-xl focus:ring-2 focus:ring-amber-500 focus:bg-white outline-none text-lg"
                    value={transferForm.newLocation}
                    onChange={(e) => setTransferForm({...transferForm, newLocation: e.target.value})}
                    placeholder="Almoxarifado Central"
                  />
                </div>
              </div>

              <div className="pt-4 border-t border-gray-100">
                <button 
                  type="submit"
                  disabled={!assets.find(a => a.barcode === transferForm.barcode)}
                  className="w-full bg-amber-500 disabled:bg-gray-200 disabled:text-gray-400 text-white px-6 py-4 rounded-xl font-bold shadow-lg shadow-amber-200 flex items-center justify-center gap-2 active:bg-amber-600 active:scale-[0.98] transition-all text-lg"
                >
                  <ArrowRightLeft size={24} /> Registrar Movimento
                </button>
              </div>
            </form>
          </div>
        )}

        {/* --- IMPORT EXCEL --- */}
        {!loading && activeTab === 'import' && (
          <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden animate-slide-in">
            <div className="bg-emerald-500/10 px-5 py-4 border-b border-emerald-100">
              <h2 className="font-bold text-xl text-emerald-900 flex items-center gap-2">
                <FileSpreadsheet size={22} className="text-emerald-600" /> Importar Planilha
              </h2>
            </div>
            
            <div className="p-5 space-y-6">
              <div className="bg-emerald-50 border border-emerald-100 rounded-xl p-4 text-sm text-emerald-900">
                <p className="font-bold mb-2">Padrão da Planilha (.xlsx):</p>
                <ol className="list-decimal pl-4 space-y-1">
                  <li>Código de Barras</li>
                  <li>Nome do Ativo</li>
                  <li>Localização</li>
                  <li>Descrição (Opcional)</li>
                </ol>
                <p className="mt-3 text-xs opacity-80">* A primeira linha (cabeçalho) é ignorada.</p>
              </div>

              <div className="w-full">
                {importProgress ? (
                  <div className="bg-gray-50 border border-gray-200 rounded-2xl p-6 text-center">
                    <div className="animate-spin rounded-full h-10 w-10 border-4 border-emerald-500 border-t-transparent mx-auto mb-4"></div>
                    <h3 className="font-bold text-gray-800 mb-3">{importProgress.status}</h3>
                    {importProgress.total > 0 && (
                      <div className="w-full bg-gray-200 rounded-full h-3 overflow-hidden">
                        <div 
                          className="bg-emerald-500 h-3 rounded-full transition-all duration-300" 
                          style={{ width: `${(importProgress.current / importProgress.total) * 100}%` }}
                        ></div>
                      </div>
                    )}
                  </div>
                ) : (
                  <label className="flex flex-col items-center justify-center w-full py-12 border-2 border-dashed border-gray-300 rounded-2xl cursor-pointer bg-gray-50 hover:bg-gray-100 active:bg-gray-200 transition-colors">
                    <UploadCloud className="w-14 h-14 mb-4 text-emerald-500" />
                    <span className="bg-emerald-600 text-white font-bold px-6 py-3 rounded-xl shadow-md">
                      Selecionar Arquivo
                    </span>
                    <p className="text-sm text-gray-500 mt-4">Somente .xlsx ou .xls</p>
                    <input 
                      type="file" 
                      className="hidden" 
                      accept=".xlsx, .xls"
                      onChange={handleExcelUpload}
                    />
                  </label>
                )}
              </div>
            </div>
          </div>
        )}

      </main>

      {/* --- BOTTOM NAVIGATION BAR --- */}
      <nav className="fixed bottom-0 left-0 right-0 bg-white border-t border-gray-200 shadow-[0_-4px_20px_rgba(0,0,0,0.05)] z-40 pb-safe">
        <div className="max-w-md mx-auto flex justify-around items-center px-2">
          <button 
            onClick={() => setActiveTab('dashboard')} 
            className={`flex flex-col items-center p-3 flex-1 transition-colors ${activeTab === 'dashboard' ? 'text-blue-600' : 'text-gray-400 hover:text-gray-600'}`}
          >
            <LayoutDashboard size={24} className={activeTab === 'dashboard' ? 'animate-bounce-slight' : ''} />
            <span className="text-[10px] font-bold mt-1">Início</span>
          </button>
          
          <button 
            onClick={() => setActiveTab('add')} 
            className={`flex flex-col items-center p-3 flex-1 transition-colors ${activeTab === 'add' ? 'text-blue-600' : 'text-gray-400 hover:text-gray-600'}`}
          >
            <PlusCircle size={24} />
            <span className="text-[10px] font-bold mt-1">Novo</span>
          </button>

          <button 
            onClick={() => setActiveTab('transfer')} 
            className={`flex flex-col items-center p-3 flex-1 transition-colors ${activeTab === 'transfer' ? 'text-amber-500' : 'text-gray-400 hover:text-gray-600'}`}
          >
            <ArrowRightLeft size={24} />
            <span className="text-[10px] font-bold mt-1">Mover</span>
          </button>

          <button 
            onClick={() => setActiveTab('import')} 
            className={`flex flex-col items-center p-3 flex-1 transition-colors ${activeTab === 'import' ? 'text-emerald-500' : 'text-gray-400 hover:text-gray-600'}`}
          >
            <FileSpreadsheet size={24} />
            <span className="text-[10px] font-bold mt-1">Importar</span>
          </button>
        </div>
      </nav>

      {/* Modals & CSS Animations */}
      {showScannerFor && <BarcodeScannerModal onScan={handleScan} onClose={() => setShowScannerFor(null)} />}
      {historyAsset && <HistoryModal asset={historyAsset} onClose={() => setHistoryAsset(null)} />}
      {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}

      <style dangerouslySetInnerHTML={{__html: `
        @keyframes fade-in-down {
          0% { opacity: 0; transform: translate(-50%, -20px); }
          100% { opacity: 1; transform: translate(-50%, 0); }
        }
        @keyframes slide-in {
          from { opacity: 0; transform: translateX(20px); }
          to { opacity: 1; transform: translateX(0); }
        }
        @keyframes slide-up {
          from { opacity: 0; transform: translateY(100%); }
          to { opacity: 1; transform: translateY(0); }
        }
        .animate-fade-in-down { animation: fade-in-down 0.3s ease-out forwards; }
        .animate-slide-in { animation: slide-in 0.3s ease-out forwards; }
        .animate-slide-up { animation: slide-up 0.3s ease-out forwards; }
        .animate-fade-in { animation: slide-in 0.2s ease-out forwards; }
        .pb-safe { padding-bottom: env(safe-area-inset-bottom); }
      `}} />
    </div>
  );
}
