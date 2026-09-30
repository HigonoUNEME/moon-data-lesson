import React, { useEffect, useRef, useState } from 'react';
import * as THREE from 'three';
import { OrbitControls } from 'three/examples/jsm/controls/OrbitControls.js';
import { LunarFeature, MoonViewerSettings } from '../types';
import { createProceduralMoonTextures, latLongToVector3, vector3ToLatLong } from '../utils/lunarTexture';
import { staticLayerValueAt } from '../utils/dataLayerLookup';
import { loadDiurnalLookup, diurnalValueAt } from '../utils/diurnalLookup';
import overlayLayers from '../data/overlayLayers.generated.json';
import diurnalFramesMeta from '../data/diurnalFrames.generated.json';
import {
  RotateCcw,
  RotateCw,
  Compass,
  Image,
  ZoomIn,
  ZoomOut,
  Play,
  Pause,
  Maximize2,
  Thermometer,
  Crosshair
} from 'lucide-react';

// requirements_v3.3 続き（「地球の風」Phase1〜3）：月面に重ねるデータ層。
// 静的な層（site_environment.csv の各指標）は webapp/react/scripts/gen_overlay_textures.py が、
// アニメーション層（1日の温度変化）は gen_diurnal_frames.py が、それぞれビルド時に生成する。
interface OverlayLayer {
  key: string; label: string; unit: string; desc: string;
  min: number; max: number; texture: string; gradientCss: string; source: string;
}
interface DiurnalFrame { index: number; subsolarLon: number; texture: string }
interface DiurnalMeta {
  nFrames: number; min: number; max: number; unit: string; label: string; desc: string;
  source: string; gradientCss: string; frames: DiurnalFrame[];
}
const DATA_LAYERS = overlayLayers as OverlayLayer[];
const DIURNAL = diurnalFramesMeta as DiurnalMeta;
const DIURNAL_KEY = 'diurnal_temp';

// 太陽は世界座標で固定し、月本体を自転させることで昼夜が移り変わるようにする（旧「太陽光照射角」
// スライダーは廃止し、自転角ひとつに統一：太陽光照射角がわかりにくいという指摘への対応）。
// この角度は sunLight の位置 (cos, 1.5, sin)*12 の式の元になった角度をそのまま定数化したもの
// （自転角0のときの見え方を、以前のデフォルト sunAngle=45 と揃えるための値）。
const FIXED_SUN_WORLD_DEG = 45;

/** 月面ローカル座標で見た「太陽直下点の経度」。太陽は世界座標で固定なので、月が rotationDeg だけ
 *  自転した分だけ月面から見た太陽の方向はズレる＝自転が進む＝昼夜が移り変わる、という対応にする。
 *  （sunLight・sunMesh の向きは変わらないまま、月面のほうが回っていく） */
function subsolarLonLocal(rotationDeg: number): number {
  const lon = -FIXED_SUN_WORLD_DEG - rotationDeg;
  return ((lon + 180) % 360 + 360) % 360 - 180;
}

/** 現在の自転角にいちばん近い、あらかじめ焼いた24フレームのうちの1枚を選ぶ。 */
function nearestDiurnalFrameIndexForRotation(rotationDeg: number): number {
  const lon = subsolarLonLocal(rotationDeg);
  const n = DIURNAL.nFrames;
  return (((Math.round((lon + 180) / (360 / n)) % n) + n) % n);
}

/** 温度系のデータ層は元データがケルビン(K)。高校生にはセルシウス度（℃）のほうが直感的という
 *  要望への対応で、画面表示だけ℃に変換する（元のCSV・PNGテクスチャ・凡例カラーマップは
 *  Kのまま。webapp/data.htmlのCSVダウンロードにも影響しない＝「ウェブアプリ上だけ」の変換）。
 *  「1日の温度差」(temp_amp_K)は絶対温度ではなく差なので、ΔK=Δ℃よりオフセットせず単位だけ変える。 */
function toCelsiusDisplay(key: string, unit: string, value: number): { value: number; unit: string } {
  if (unit !== 'K') return { value, unit };
  const isDelta = key === 'temp_amp_K';
  return { value: isDelta ? value : value - 273.15, unit: '℃' };
}

// 常設ピンを廃止し、カーソルを合わせたときだけ「近くの地点」の名前を出す（要望への対応）。
// これより離れていたら「地点なし」＝緯度経度だけを表示する。
const FEATURE_HOVER_THRESHOLD_DEG = 3;

/** features の中で lat, lon にいちばん近いものを探す。閾値より遠ければ null。
 *  緯度によって経度1度あたりの実距離が縮む分を cos(lat) で補正した簡易距離（moonkit.nearest() と同じ考え方）。*/
function nearestFeature(features: LunarFeature[], lat: number, lon: number): LunarFeature | null {
  let best: LunarFeature | null = null;
  let bestD = Infinity;
  const cosLat = Math.cos((lat * Math.PI) / 180);
  for (const f of features) {
    const dLat = f.latitude - lat;
    let dLon = f.longitude - lon;
    if (dLon > 180) dLon -= 360;
    if (dLon < -180) dLon += 360;
    dLon *= cosLat;
    const d = dLat * dLat + dLon * dLon;
    if (d < bestD) {
      bestD = d;
      best = f;
    }
  }
  return best && Math.sqrt(bestD) <= FEATURE_HOVER_THRESHOLD_DEG ? best : null;
}

const CATEGORY_EMOJI: Record<string, string> = {
  missions: '🚀 着陸・探査機ミッション',
  craters: '🌕 衝突クレーター',
  maria: '🌊 月の海・大盆地',
  moonquakes: '⚡ 月震・観測イベント',
  resources: '💧 極域資源・水氷候補'
};

/** カーソルを合わせた地点の情報。常設ピンの代わりにホバーで出す（要望への対応）。 */
interface HoverInfo {
  lat: number;
  lon: number;
  feature: LunarFeature | null;               // 閾値内に既知の地点があれば
  layerValue: { label: string; unit: string; value: number } | null; // 重ねているデータ層の値
}

interface MoonViewer3DProps {
  /** 名前表示・クリック選択のためのホバー判定に使う全地点（データセットの絞り込みとは無関係）。 */
  features: LunarFeature[];
  selectedFeature: LunarFeature | null;
  onSelectFeature: (feature: LunarFeature) => void;
  settings: MoonViewerSettings;
  onUpdateSettings: (settings: Partial<MoonViewerSettings>) => void;
  /** true のとき、左ペインいっぱいに広げる（角丸カードにしない） */
  fill?: boolean;
}

const MOON_RADIUS = 2.0;
const START_DISTANCE = 6.4;
const FOCUS_DISTANCE = 4.8;
const MIN_DISTANCE = 3.0;
const MAX_DISTANCE = 12;

// 遊び心：地球と太陽を「実際の見え方」で置く。実際の半径・距離の比から視半径（見かけの角度）を
// 計算し、月から十分離れた仮想の球殻（半径 SKY_R）の上に、その角度どおりの大きさで置く
// （mesh半径 = SKY_R * tan(視半径)。SKY_R をどう選んでも見かけの角度は変わらない）。
const SKY_R = 120;
const R_EARTH_KM = 6371;
const R_SUN_KM = 696000;
const D_EARTH_MOON_KM = 384400;      // 地球〜月の平均距離
const D_EARTH_SUN_KM = 149_600_000;  // 1天文単位（太陽は月からも地球からもほぼ同じ角度に見える）
const EARTH_ANGULAR_R = Math.atan(R_EARTH_KM / D_EARTH_MOON_KM);
const SUN_ANGULAR_R = Math.atan(R_SUN_KM / D_EARTH_SUN_KM);
const EARTH_MESH_R = SKY_R * Math.tan(EARTH_ANGULAR_R);
const SUN_MESH_R = SKY_R * Math.tan(SUN_ANGULAR_R);
const Y_AXIS = new THREE.Vector3(0, 1, 0);
// 月面の自転速度（見て楽しいよう誇張した速さ。実際の公転・自転周期＝約27.3日を再現する意図ではない）
const SPIN_DEG_PER_SEC_AT_1X = 6;

/** 太陽の周りにふわっと光る輪（グロー）のテクスチャを作る。実際の球体の大きさ（=見た目の視半径）は
 *  変えず、写真で太陽が眩しく滲んで写るのと同じ効果を加算合成で足すことで見つけやすくする
 *  （「実際の大きさ」を誇張せずに目立たせるための表現）。 */
function createSunGlowTexture(): THREE.CanvasTexture {
  const size = 128;
  const canvas = document.createElement('canvas');
  canvas.width = size;
  canvas.height = size;
  const ctx = canvas.getContext('2d')!;
  const grad = ctx.createRadialGradient(size / 2, size / 2, 0, size / 2, size / 2, size / 2);
  grad.addColorStop(0, 'rgba(255,247,224,0.95)');
  grad.addColorStop(0.25, 'rgba(255,230,170,0.55)');
  grad.addColorStop(0.6, 'rgba(255,200,120,0.16)');
  grad.addColorStop(1, 'rgba(255,200,120,0)');
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, size, size);
  return new THREE.CanvasTexture(canvas);
}

// 月面テクスチャの解像度段階。初期表示は軽い2Kのまま、ズームしたときだけ重い方を読みに行く
// （NASA SVS「CGI Moon Kit」2019年版。webapp/react/scripts/build_moon_textures.py が用意）。
type MoonTexTier = '2k' | '4k' | '8k';
const MOON_TEX_URL: Record<MoonTexTier, string> = {
  '2k': 'textures/moon_lroc_color_2k.jpg',
  '4k': 'textures/moon_lroc_color_4k.jpg',
  '8k': 'textures/moon_lroc_color_8k.jpg',
};
/** カメラが月の中心からどれだけ離れているかで、使うテクスチャの段階を決める。
 *  FOCUS_DISTANCE(=4.8, 地点を選んだときの寄り) で4Kに、手動でさらに寄る(MIN_DISTANCE寄り)と8Kになる。 */
function pickMoonTexTier(distance: number): MoonTexTier {
  if (distance <= 4.0) return '8k';
  if (distance <= 5.5) return '4k';
  return '2k';
}

export const MoonViewer3D: React.FC<MoonViewer3DProps> = ({
  features,
  selectedFeature,
  onSelectFeature,
  settings,
  onUpdateSettings,
  fill = false
}) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);

  // Three.js instances
  const sceneRef = useRef<THREE.Scene | null>(null);
  const cameraRef = useRef<THREE.PerspectiveCamera | null>(null);
  const rendererRef = useRef<THREE.WebGLRenderer | null>(null);
  const controlsRef = useRef<OrbitControls | null>(null);
  const moonMeshRef = useRef<THREE.Mesh | null>(null);
  const moonSpinGroupRef = useRef<THREE.Group | null>(null); // 自転で回す本体（月面・データ層・グリッド・ピン）
  const overlayMeshRef = useRef<THREE.Mesh | null>(null);
  const overlayMaterialRef = useRef<THREE.MeshBasicMaterial | null>(null);
  const overlayGridGroupRef = useRef<THREE.Group | null>(null);
  const sunLightRef = useRef<THREE.DirectionalLight | null>(null);
  const earthMeshRef = useRef<THREE.Mesh | null>(null); // 遊び心の地球（潮汐固定＝自転と同じ速さで空を巡る）
  const sunMeshRef = useRef<THREE.Mesh | null>(null);    // 遊び心の太陽（世界座標で固定。月が自転する）
  const animationFrameIdRef = useRef<number | null>(null);

  // データ層のテクスチャキャッシュ（url -> 読み込み Promise）。切り替えるたびに読み直さないため
  const textureCacheRef = useRef<Map<string, Promise<THREE.Texture>>>(new Map());
  // 月面写真の現在の解像度段階（ズームで 2k → 4k → 8k と上げる。下げてキャッシュ済みなら再読込なし）
  const moonTexTierRef = useRef<MoonTexTier>('2k');
  // 現在読み込み済みの月面実写テクスチャ（表示オフ中もここに保持し、オンに戻したら再適用する）
  const moonPhotoTextureRef = useRef<THREE.Texture | null>(null);
  // 1日の温度アニメーション（24枚）。選んだときに一度だけ読み込む
  const diurnalTexturesRef = useRef<THREE.Texture[] | null>(null);
  const [diurnalLoading, setDiurnalLoading] = useState(false);
  // 1日の温度アニメーションの「ホバー時の値」用ルックアップ（3度グリッド）。選んだときに一度だけ読み込む
  const diurnalLookupRef = useRef<{ step: number; cols: string[]; cells: number[][] } | null>(null);

  // 選択地点へカメラを寄せるためのトゥイーン目標（azimuth / polar / distance）。null のとき何もしない
  const focusRef = useRef<{ az: number; pol: number; dist: number } | null>(null);
  // クリックとドラッグを区別するためのポインタ押下位置
  const pointerDownRef = useRef<{ x: number; y: number; t: number } | null>(null);
  // 自転中もマウス直下の地点を追従させるため、最後にポインタがあった画面座標を animate ループから読む
  const lastPointerClientRef = useRef<{ x: number; y: number } | null>(null);
  // クリックして立てたピンの情報。animate ループ（クロージャ）から最新値を読むための ref
  const pinnedInfoRef = useRef<HoverInfo | null>(null);
  const pinMarkerRef = useRef<THREE.Group | null>(null); // クリックした地点に立てる目印（moonSpinGroup の子）
  // 自転していてもピンを見失わないようにする「追従」の on/off。animate ループから読むための ref も持つ
  const followPinRef = useRef(false);

  // HUD & Hover state。常設ピンは廃止したので、カーソルを合わせた地点の情報をまとめて持つ
  const [hoverInfo, setHoverInfo] = useState<HoverInfo | null>(null);
  const [tooltipPos, setTooltipPos] = useState<{ x: number; y: number } | null>(null);
  // クリックして立てたピンの情報（ホバーと違い、マウスを離しても消えない）
  const [pinnedInfo, setPinnedInfo] = useState<HoverInfo | null>(null);
  // 自転してもピンを見失わないよう、カメラの向き（経度方向だけ）をピンに合わせ続けるか
  const [followPin, setFollowPin] = useState(false);
  followPinRef.current = followPin;
  // 左下の展開図に出す「今カメラが向いている地点」（月面ローカルの緯度経度）＋見えている範囲の半径
  // （limbDeg。カメラが有限距離にあるので、90°の半球まるごとではなく acos(R/カメラ距離) までしか
  // 実際には見えない。「青丸1点しか見えていないわけではない」という指摘への対応で、点ではなく
  // 範囲として展開図に描く）。animate ループから更新。
  const [viewCenter, setViewCenter] = useState<{ lat: number; lon: number; limbDeg: number }>({ lat: 0, lon: 0, limbDeg: 70 });
  const viewCenterRef = useRef<{ lat: number; lon: number; limbDeg: number }>({ lat: 0, lon: 0, limbDeg: 70 });
  const miniMapCanvasRef = useRef<HTMLCanvasElement>(null); // 展開図：見えている範囲を淡く塗る

  // settings を animate ループの外（クロージャ）から読むための ref
  const settingsRef = useRef(settings);
  settingsRef.current = settings;
  // 自転アニメーション中、スライダー表示を追従させるために animate ループから呼ぶ（常に最新を指す）
  const onUpdateSettingsRef = useRef(onUpdateSettings);
  onUpdateSettingsRef.current = onUpdateSettings;

  // features は props なので毎レンダー変わりうる。animate ループ（1回だけ張るクロージャ）から
  // 常に最新を読めるよう ref に写す（settingsRef と同じ考え方）。
  const featuresRef = useRef(features);
  featuresRef.current = features;

  /** 月面ローカルの緯度経度における、現在のデータ層の値を求める（表示中でなければ null）。
   *  ホバー表示・クリックで立てたピンの両方から、同じ計算式で使う。 */
  const computeLayerValueAt = (lat: number, lon: number): HoverInfo['layerValue'] => {
    const s = settingsRef.current;
    if (!s.showDataLayer) return null;
    if (s.dataLayerKey === DIURNAL_KEY) {
      const lookup = diurnalLookupRef.current;
      if (!lookup) return null;
      const rotationDeg = moonSpinGroupRef.current
        ? THREE.MathUtils.radToDeg(moonSpinGroupRef.current.rotation.y)
        : s.moonRotationDeg;
      const v = diurnalValueAt(lookup, lat, lon, subsolarLonLocal(rotationDeg));
      if (v === null) return null;
      const disp = toCelsiusDisplay(DIURNAL_KEY, DIURNAL.unit, v);
      return { label: DIURNAL.label, unit: disp.unit, value: disp.value };
    }
    const v = staticLayerValueAt(s.dataLayerKey, lat, lon);
    if (v === null) return null;
    const layer = DATA_LAYERS.find((l) => l.key === s.dataLayerKey);
    if (!layer) return null;
    const disp = toCelsiusDisplay(layer.key, layer.unit, v);
    return { label: layer.label, unit: disp.unit, value: disp.value };
  };

  /** 画面座標(clientX/Y)から、月面と交わった点の緯度経度・近くの既知地点・データ層の値をまとめて求める。
   *  ホバー表示（自転に追従させるため毎フレーム）とクリック選択・ピン設置の両方から使う共通ロジック。
   *  animate ループより前に定義することで、そちらのクロージャからも直接参照できるようにしている。 */
  const getHoverInfoAtClient = (clientX: number, clientY: number): HoverInfo | null => {
    const rect = canvasRef.current?.getBoundingClientRect();
    if (!rect || !cameraRef.current || !moonMeshRef.current) return null;
    const ndc = new THREE.Vector2(
      ((clientX - rect.left) / rect.width) * 2 - 1,
      -((clientY - rect.top) / rect.height) * 2 + 1
    );
    const rc = new THREE.Raycaster();
    rc.setFromCamera(ndc, cameraRef.current);
    const hits = rc.intersectObject(moonMeshRef.current);
    if (!hits.length) return null;
    const localPoint = moonSpinGroupRef.current
      ? moonSpinGroupRef.current.worldToLocal(hits[0].point.clone())
      : hits[0].point;
    const { lat, lon } = vector3ToLatLong(localPoint);
    const feature = nearestFeature(featuresRef.current, lat, lon);
    const layerValue = computeLayerValueAt(lat, lon);
    return { lat: +lat.toFixed(2), lon: +lon.toFixed(2), feature, layerValue };
  };

  /** データ層用テクスチャを読み込む（キャッシュ付き）。同じ URL を2回読みに行かない。 */
  const loadTexture = (url: string): Promise<THREE.Texture> => {
    const cache = textureCacheRef.current;
    let p = cache.get(url);
    if (!p) {
      p = new Promise<THREE.Texture>((resolve, reject) => {
        new THREE.TextureLoader().load(
          url,
          (tex) => {
            tex.wrapS = THREE.RepeatWrapping;
            tex.wrapT = THREE.ClampToEdgeWrapping;
            resolve(tex);
          },
          undefined,
          reject
        );
      });
      cache.set(url, p);
    }
    return p;
  };

  /** 月の実写テクスチャ（色）を表示するかどうかを、今の月面メッシュに反映する。
   *  「データ層の色なのか、月面写真自体の色なのか分からない」という指摘への対応：
   *  オフのときは色テクスチャだけ外し（bumpMapは残すので陰影自体は見える）、素の灰色にする。
   *  animate ループより前に定義し、テクスチャの読み込み完了時・設定変更時の両方から呼べるようにする。 */
  const applyMoonTextureVisibility = () => {
    const mat = moonMeshRef.current?.material;
    if (!(mat instanceof THREE.MeshStandardMaterial)) return;
    if (settingsRef.current.showMoonTexture) {
      mat.map = moonPhotoTextureRef.current;
      mat.color.setHex(0xffffff);
    } else {
      mat.map = null;
      mat.color.setHex(0x8a8a8a);
    }
    mat.needsUpdate = true;
  };

  // Initialize Three.js Scene (once)
  useEffect(() => {
    if (!containerRef.current || !canvasRef.current) return;

    const width = containerRef.current.clientWidth || 800;
    const height = containerRef.current.clientHeight || 480;

    const scene = new THREE.Scene();
    sceneRef.current = scene;

    // far は星空(半径60〜140)・太陽/地球(SKY_R=120)が入る範囲で十分な深さに絞る。
    // 大きすぎる(旧2000)と遠くの物体の奥行き精度が足りず、太陽・地球がまれに描画から
    // 欠落する（深度バッファの精度不足でクリップされる）ことがあったための修正。
    const camera = new THREE.PerspectiveCamera(42, width / height, 0.1, 300);
    camera.position.set(2.2, 1.4, START_DISTANCE);
    cameraRef.current = camera;

    const renderer = new THREE.WebGLRenderer({
      canvas: canvasRef.current,
      antialias: true,
      alpha: true,
      powerPreference: 'high-performance'
    });
    renderer.setSize(width, height);
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    renderer.toneMapping = THREE.ACESFilmicToneMapping;
    renderer.toneMappingExposure = 1.15;
    rendererRef.current = renderer;

    // --- OrbitControls：標準の軌道カメラ。ドラッグで回転・ホイールで拡大縮小・慣性つき ---
    const controls = new OrbitControls(camera, renderer.domElement);
    controls.target.set(0, 0, 0);
    controls.enableDamping = true;
    controls.dampingFactor = 0.08;
    controls.enablePan = false;              // 月儀なので平行移動は無効
    controls.rotateSpeed = 0.55;
    controls.zoomSpeed = 0.9;
    controls.minDistance = MIN_DISTANCE;
    controls.maxDistance = MAX_DISTANCE;
    // 「自動回転」は OrbitControls のカメラ回転ではなく、月本体（moonSpinGroup）を回す方式にした
    // （下の animate ループ参照）。カメラ自体は常にユーザー操作待ち。
    controls.update();
    controls.saveState();                    // reset() の戻り先
    controlsRef.current = controls;

    // Starfield
    const starsGeo = new THREE.BufferGeometry();
    const starCount = 1400;
    const starPositions = new Float32Array(starCount * 3);
    const starColors = new Float32Array(starCount * 3);
    for (let i = 0; i < starCount * 3; i += 3) {
      const radius = 60 + Math.random() * 80;
      const u = Math.random();
      const v = Math.random();
      const theta = u * 2.0 * Math.PI;
      const phi = Math.acos(2.0 * v - 1.0);
      const sinPhi = Math.sin(phi);
      starPositions[i] = radius * sinPhi * Math.cos(theta);
      starPositions[i + 1] = radius * sinPhi * Math.sin(theta);
      starPositions[i + 2] = radius * Math.cos(phi);
      const shade = 0.8 + Math.random() * 0.2;
      starColors[i] = shade;
      starColors[i + 1] = shade * (0.9 + Math.random() * 0.1);
      starColors[i + 2] = shade * (0.95 + Math.random() * 0.05);
    }
    starsGeo.setAttribute('position', new THREE.BufferAttribute(starPositions, 3));
    starsGeo.setAttribute('color', new THREE.BufferAttribute(starColors, 3));
    scene.add(new THREE.Points(starsGeo, new THREE.PointsMaterial({
      size: 1.2, vertexColors: true, transparent: true, opacity: 0.85
    })));

    // 自転で回す本体グループ。月面写真・データ層・グリッドはすべてこの中に入れ、
    // まとめて rotation.y を回すことで「カメラではなく月自身が自転する」ようにする。
    const moonSpinGroup = new THREE.Group();
    moonSpinGroup.rotation.y = THREE.MathUtils.degToRad(settings.moonRotationDeg);
    scene.add(moonSpinGroup);
    moonSpinGroupRef.current = moonSpinGroup;

    // Moon globe（手続き生成テクスチャ → 実写に差し替え）
    const { colorTexture, bumpTexture } = createProceduralMoonTextures();
    const moonMaterial = new THREE.MeshStandardMaterial({
      map: colorTexture, bumpMap: bumpTexture, bumpScale: 0.045,
      roughness: 0.92, metalness: 0.05, wireframe: settings.wireframe
    });
    const moonMesh = new THREE.Mesh(new THREE.SphereGeometry(MOON_RADIUS, 64, 64), moonMaterial);
    moonSpinGroup.add(moonMesh);
    moonMeshRef.current = moonMesh;

    let photoTexture: THREE.Texture | null = null;
    new THREE.TextureLoader().load(
      'textures/moon_lroc_color_2k.jpg',
      (tex) => {
        tex.colorSpace = THREE.SRGBColorSpace;
        tex.wrapS = THREE.RepeatWrapping;
        tex.wrapT = THREE.ClampToEdgeWrapping;
        tex.anisotropy = renderer.capabilities.getMaxAnisotropy();
        photoTexture = tex;
        moonPhotoTextureRef.current = tex;
        if (moonMeshRef.current?.material instanceof THREE.MeshStandardMaterial) {
          const m = moonMeshRef.current.material;
          m.bumpMap = tex; // 陰影用。色テクスチャの表示オン/オフに関わらず常に使う
          m.bumpScale = 0.012;
          applyMoonTextureVisibility();
        }
      },
      undefined,
      () => console.warn('月面テクスチャの読み込みに失敗。手続き生成テクスチャを使用します。')
    );

    // データ層（半透明の外殻）。requirements_v3.3 続き（「地球の風」Phase1〜3）。
    // 月本体のすぐ外側にもう1枚球を重ね、生成済みのカラーマップ画像を貼る。
    // どの画像を貼るか（静的な層／1日のアニメーション）は、別の useEffect（layer切り替え）が担当する。
    const overlayMaterial = new THREE.MeshBasicMaterial({
      transparent: true, opacity: 0.8, depthWrite: false, toneMapped: false
    });
    overlayMaterialRef.current = overlayMaterial;
    const overlayMesh = new THREE.Mesh(
      new THREE.SphereGeometry(MOON_RADIUS + 0.01, 64, 64), overlayMaterial);
    overlayMesh.renderOrder = 1; // 月本体の後に描く（半透明の重ね順を安定させる）
    overlayMesh.visible = settings.showDataLayer; // 表示・非表示は mesh 側で管理する
    moonSpinGroup.add(overlayMesh);
    overlayMeshRef.current = overlayMesh;

    // データ層の上に重ねる15度おきの白いグリッド（「地球の風」nullschool.net を参考にした読み取り補助線）
    const overlayGridGroup = new THREE.Group();
    const graticuleMat = new THREE.LineBasicMaterial({ color: 0xffffff, transparent: true, opacity: 0.28 });
    const GRID_R = MOON_RADIUS + 0.012;
    for (let lat = -75; lat <= 75; lat += 15) {
      const pts: THREE.Vector3[] = [];
      for (let j = 0; j <= 72; j++) pts.push(latLongToVector3(lat, (j / 72) * 360 - 180, GRID_R));
      overlayGridGroup.add(new THREE.Line(new THREE.BufferGeometry().setFromPoints(pts), graticuleMat));
    }
    for (let lon = -180; lon < 180; lon += 15) {
      const pts: THREE.Vector3[] = [];
      for (let j = 0; j <= 36; j++) pts.push(latLongToVector3(-90 + (j / 36) * 180, lon, GRID_R));
      overlayGridGroup.add(new THREE.Line(new THREE.BufferGeometry().setFromPoints(pts), graticuleMat));
    }
    overlayGridGroup.renderOrder = 2;
    overlayGridGroup.visible = settings.showDataLayer;
    moonSpinGroup.add(overlayGridGroup);
    overlayGridGroupRef.current = overlayGridGroup;

    // Coordinate grid
    const gridGroup = new THREE.Group();
    const equatorPoints: THREE.Vector3[] = [];
    for (let i = 0; i <= 64; i++) {
      const a = (i / 64) * Math.PI * 2;
      equatorPoints.push(new THREE.Vector3(
        (MOON_RADIUS + 0.005) * Math.cos(a), 0, (MOON_RADIUS + 0.005) * Math.sin(a)));
    }
    gridGroup.add(new THREE.Line(
      new THREE.BufferGeometry().setFromPoints(equatorPoints),
      new THREE.LineBasicMaterial({ color: 0x00f0ff, transparent: true, opacity: 0.45 })));

    const meridianPoints: THREE.Vector3[] = [];
    for (let i = 0; i <= 64; i++) {
      meridianPoints.push(latLongToVector3(-90 + (i / 64) * 180, 0, MOON_RADIUS + 0.005));
    }
    gridGroup.add(new THREE.Line(
      new THREE.BufferGeometry().setFromPoints(meridianPoints),
      new THREE.LineBasicMaterial({ color: 0xffb703, transparent: true, opacity: 0.45 })));

    [-60, -30, 30, 60].forEach(latDeg => {
      const pts: THREE.Vector3[] = [];
      for (let j = 0; j <= 64; j++) pts.push(latLongToVector3(latDeg, (j / 64) * 360 - 180, MOON_RADIUS + 0.003));
      gridGroup.add(new THREE.Line(
        new THREE.BufferGeometry().setFromPoints(pts),
        new THREE.LineBasicMaterial({ color: 0x64748b, transparent: true, opacity: 0.25 })));
    });

    const axisLine = new THREE.Line(
      new THREE.BufferGeometry().setFromPoints([
        new THREE.Vector3(0, MOON_RADIUS + 0.35, 0),
        new THREE.Vector3(0, -MOON_RADIUS - 0.35, 0)]),
      new THREE.LineDashedMaterial({
        color: 0x94a3b8, dashSize: 0.08, gapSize: 0.04, transparent: true, opacity: 0.5 }));
    axisLine.computeLineDistances();
    gridGroup.add(axisLine);

    // 緯度経度グリッドは常時表示（トグル廃止。月画像の表示切替で代わりに見やすさを確保する）
    moonSpinGroup.add(gridGroup);

    // クリックした地点に立てる目印（要望：「クリックしてある地点にピンを立てて、データを表示する」）。
    // moonSpinGroup の子にすることで、自転しても地点に張り付いたまま一緒に回る。
    const pinGroup = new THREE.Group();
    const pinColor = 0xfacc15;
    const pinStickH = 0.16;
    const pinNeedle = new THREE.Mesh(
      new THREE.CylinderGeometry(0.012, 0.012, pinStickH, 8),
      new THREE.MeshBasicMaterial({ color: pinColor, toneMapped: false })
    );
    pinNeedle.position.y = pinStickH / 2;
    const pinHead = new THREE.Mesh(
      new THREE.SphereGeometry(0.045, 14, 14),
      new THREE.MeshBasicMaterial({ color: pinColor, toneMapped: false })
    );
    pinHead.position.y = pinStickH;
    pinGroup.add(pinNeedle, pinHead);
    pinGroup.renderOrder = 5; // データ層・グリッドより手前に描く
    pinGroup.visible = false;
    moonSpinGroup.add(pinGroup);
    pinMarkerRef.current = pinGroup;

    // Lights（太陽は世界座標で固定。動くのは月本体のほう＝自転で昼夜が移り変わる）
    const fixedSunRad = (FIXED_SUN_WORLD_DEG * Math.PI) / 180;
    const fixedSunDir = new THREE.Vector3(Math.cos(fixedSunRad) * 12, 1.5, Math.sin(fixedSunRad) * 12);
    const sunLight = new THREE.DirectionalLight(0xfff8f0, settings.lightIntensity);
    sunLight.position.copy(fixedSunDir);
    scene.add(sunLight);
    sunLightRef.current = sunLight;

    // 遊び心：地球と太陽を実際の大きさ・見え方で置く（webapp/react/scripts/build_planet_textures.py で用意）。
    // 地球は月から見て潮汐固定＝常に同じ面を向けているので、月面の「表側中心」（緯度0・経度0）が
    // 向く方向に固定し、自転アニメーションでは月本体と同じ角速度で一緒に空を巡らせる（下の animate 参照）。
    // 光源の向き次第で暗く沈んで見えなくなってしまうため、太陽と同じく自ら光って見える
    // MeshBasicMaterial にする（実際のアポロ「地球の出」写真のような、常によく見える地球にする簡略化）。
    let earthTexture: THREE.Texture | null = null;
    // MeshBasicMaterialのcolorはmapに乗算されるので、地球をもう少し明るくという要望への対応で
    // 元の 0x6f93c4 より各チャンネルを底上げ（青みがかった色味は保ちつつ、暗く沈みすぎないように）。
    const earthMesh = new THREE.Mesh(
      new THREE.SphereGeometry(EARTH_MESH_R, 48, 48),
      new THREE.MeshBasicMaterial({ color: 0x8fb3e0, toneMapped: false })
    );
    earthMesh.position.copy(latLongToVector3(0, 0, SKY_R));
    scene.add(earthMesh);
    earthMeshRef.current = earthMesh;
    new THREE.TextureLoader().load(
      'textures/earth_daymap_2k.jpg',
      (tex) => {
        tex.colorSpace = THREE.SRGBColorSpace;
        earthTexture = tex;
        if (earthMesh.material instanceof THREE.MeshBasicMaterial) {
          earthMesh.material.map = tex;
          earthMesh.material.needsUpdate = true;
        }
      },
      undefined,
      () => console.warn('地球テクスチャの読み込みに失敗しました。')
    );

    // 太陽の見た目（球）も、光源と同じ固定方向に置く。自ら光るのでライトの影響を受けない
    // MeshBasicMaterial を使う。
    let sunTexture: THREE.Texture | null = null;
    const sunMesh = new THREE.Mesh(
      new THREE.SphereGeometry(SUN_MESH_R, 32, 32),
      new THREE.MeshBasicMaterial({ color: 0xfff4d6, toneMapped: false })
    );
    sunMesh.position.copy(fixedSunDir.clone().normalize().multiplyScalar(SKY_R));
    scene.add(sunMesh);
    sunMeshRef.current = sunMesh;
    new THREE.TextureLoader().load(
      'textures/sun_2k.jpg',
      (tex) => {
        tex.colorSpace = THREE.SRGBColorSpace;
        sunTexture = tex;
        if (sunMesh.material instanceof THREE.MeshBasicMaterial) {
          sunMesh.material.map = tex;
          sunMesh.material.needsUpdate = true;
        }
      },
      undefined,
      () => console.warn('太陽テクスチャの読み込みに失敗しました。')
    );

    // 太陽まわりのグロー（太陽をもう少し目立たせたいという要望への対応）。実際の球体の大きさは
    // そのまま、加算合成のふわっとした光の輪を一回り大きく重ねるだけ＝実際の見かけの大きさを
    // 誇張せずに見つけやすくする。常にカメラを向く Sprite なので、球体と違って必ず丸く見える。
    const sunGlowTexture = createSunGlowTexture();
    const sunGlow = new THREE.Sprite(new THREE.SpriteMaterial({
      map: sunGlowTexture, color: 0xfff2c8, transparent: true, depthWrite: false,
      blending: THREE.AdditiveBlending, toneMapped: false
    }));
    sunGlow.scale.setScalar(SUN_MESH_R * 6);
    sunGlow.position.copy(sunMesh.position);
    sunGlow.renderOrder = 1;
    scene.add(sunGlow);

    scene.add(new THREE.AmbientLight(0x2a3040, 0.9));
    const rimLight = new THREE.DirectionalLight(0x38bdf8, 0.35);
    rimLight.position.set(-10, -5, -8);
    scene.add(rimLight);
    // カメラに付く弱い補助光。自動回転で夜側に回っても手前の面が真っ暗にならないように
    const headLight = new THREE.DirectionalLight(0xdfe8ff, 0.75);
    camera.add(headLight);
    headLight.position.set(0, 0, 1);
    scene.add(camera);

    // Render loop
    const clock = new THREE.Clock();
    const spherical = new THREE.Spherical();
    let lastT = 0;
    let rotationSyncAccum = 0; // スライダー表示への書き戻しを間引くための積算時間
    let hoverSyncAccum = 0; // ホバー・ピンのデータ再計算を間引くための積算時間
    const earthBaseDir = latLongToVector3(0, 0, SKY_R); // 潮汐固定＝月面「表側中心」が向く方向（自転角0のときの地球の位置）

    const animate = () => {
      animationFrameIdRef.current = requestAnimationFrame(animate);
      const t = clock.getElapsedTime();
      const dt = Math.min(0.1, t - lastT);
      lastT = t;

      // 自転（月本体を回す。カメラは動かさない）。地球は潮汐固定なので、月と同じ角速度で
      // 「表側中心」の向く方向＝空の位置も一緒に回す（自転と公転が連動していることの再現）。
      // 選択地点へのトゥイーン中は、寄せた視点がずれないよう自転も止める。
      const focus0 = focusRef.current;
      if (settingsRef.current.autoRotate && !focus0) {
        // rotationSpeed はスライダーで 0.1〜3 の範囲に直接調整できるので、ここでの下駄（旧 0.3）は外す。
        const spinRad = THREE.MathUtils.degToRad(SPIN_DEG_PER_SEC_AT_1X * Math.max(0.05, settingsRef.current.rotationSpeed)) * dt;
        moonSpinGroup.rotation.y += spinRad;

        // スライダー表示（自転角）を追従させる。毎フレームだと重いので間引く
        rotationSyncAccum += dt;
        if (rotationSyncAccum > 0.12) {
          rotationSyncAccum = 0;
          const deg = ((THREE.MathUtils.radToDeg(moonSpinGroup.rotation.y) % 360) + 360) % 360;
          onUpdateSettingsRef.current({ moonRotationDeg: Math.round(deg * 10) / 10 });
        }
      }
      earthMesh.position.copy(earthBaseDir).applyAxisAngle(Y_AXIS, moonSpinGroup.rotation.y);

      // データ層が「1日の温度アニメーション」のとき、今の自転角に対応するフレームに直す。
      // 太陽は世界座標で固定なので、自転が進む＝月面から見た太陽の方向が変わる＝昼夜が移り変わる。
      // （手動スライダーでの変更は下の useEffect が rotation.y に反映するので、ここで両方まかなえる）
      if (settingsRef.current.dataLayerKey === DIURNAL_KEY && diurnalTexturesRef.current) {
        const rotationDeg = THREE.MathUtils.radToDeg(moonSpinGroup.rotation.y);
        const idx = nearestDiurnalFrameIndexForRotation(rotationDeg);
        const tex = diurnalTexturesRef.current[idx];
        if (tex && overlayMaterial.map !== tex) {
          overlayMaterial.map = tex;
          overlayMaterial.needsUpdate = true;
        }
      }

      // 選択地点へゆっくり寄せる（トゥイーン中は自動回転を止める）
      const focus = focusRef.current;
      if (focus) {
        spherical.setFromVector3(camera.position.clone().sub(controls.target));
        // 角度は最短方向で補間
        let dAz = focus.az - spherical.theta;
        while (dAz > Math.PI) dAz -= Math.PI * 2;
        while (dAz < -Math.PI) dAz += Math.PI * 2;
        spherical.theta += dAz * 0.12;
        spherical.phi += (focus.pol - spherical.phi) * 0.12;
        spherical.radius += (focus.dist - spherical.radius) * 0.12;
        spherical.makeSafe();
        camera.position.setFromSpherical(spherical).add(controls.target);
        if (Math.abs(dAz) < 0.005 &&
            Math.abs(focus.pol - spherical.phi) < 0.005 &&
            Math.abs(focus.dist - spherical.radius) < 0.02) {
          focusRef.current = null;
        }
      }

      // ピンへの「追従」：自転で経度方向にだけ動いていくピンを見失わないよう、カメラの
      // 経度（方位角 theta）だけをピンの現在位置に合わせ続ける。緯度に当たる極角（phi）と
      // ズーム（radius）はユーザーの操作をそのまま尊重する（自転は Y 軸まわりだけなので、
      // ピンの緯度＝phi は自転では変化しない。ずれるのは経度＝theta だけ）。
      if (!focus && followPinRef.current && pinMarkerRef.current?.visible) {
        // ピンの方位角は自転でずっと動き続ける「動く的」なので、選択地点フォーカスのときのような
        // 減衰補間（毎フレーム差の一部だけ詰める）だと、動き続ける的に対して恒常的に遅れてしまう
        // （速度に比例したズレが残り続ける）。ここは毎フレームそのまま合わせて、ズレを残さない。
        const pinWorld = pinMarkerRef.current.getWorldPosition(new THREE.Vector3());
        const pinSph = new THREE.Spherical().setFromVector3(pinWorld);
        spherical.setFromVector3(camera.position.clone().sub(controls.target));
        spherical.theta = pinSph.theta;
        spherical.makeSafe();
        camera.position.setFromSpherical(spherical).add(controls.target);
      }

      // マウスが止まっていても、月本体のほうが自転で動くので、直下の地点は変わり続ける。
      // pointermove イベントだけに頼ると自転中は表示が古いままになる（「ついてこない」指摘への対応）。
      // 間引きながら毎フレーム再計算し、ホバー中のカーソル位置・ピン留めした地点の両方を追従させる。
      hoverSyncAccum += dt;
      if (hoverSyncAccum > 0.08) {
        hoverSyncAccum = 0;
        const lp = lastPointerClientRef.current;
        if (lp) {
          const info = getHoverInfoAtClient(lp.x, lp.y);
          setHoverInfo((prev) => {
            if (!info) return prev === null ? prev : null;
            if (prev && prev.lat === info.lat && prev.lon === info.lon &&
                prev.feature === info.feature &&
                prev.layerValue?.value === info.layerValue?.value) return prev;
            return info;
          });
        }
        const pinned = pinnedInfoRef.current;
        if (pinned) {
          const lv = computeLayerValueAt(pinned.lat, pinned.lon);
          if (lv?.value !== pinned.layerValue?.value || lv?.label !== pinned.layerValue?.label) {
            const updated = { ...pinned, layerValue: lv };
            pinnedInfoRef.current = updated;
            setPinnedInfo(updated);
          }
        }

        // 左下の展開図用：今カメラが向いている地点（月面ローカルの緯度経度）と、
        // 実際に見えている範囲の半径（limbDeg）。カメラの位置を moonSpinGroup のローカル座標に
        // 戻す＝自転を打ち消した「向き」になる。limbDeg はカメラが有限距離にあることによる
        // 見かけの地平線の角度（acos(半径/距離)。近づくほど狭く、離れるほど90°に近づく）。
        const camLocal = moonSpinGroup.worldToLocal(camera.position.clone());
        const vc = vector3ToLatLong(camLocal);
        const camDistNow = camera.position.distanceTo(controls.target);
        const limbDeg = THREE.MathUtils.radToDeg(Math.acos(THREE.MathUtils.clamp(MOON_RADIUS / camDistNow, -1, 1)));
        const rvc = {
          lat: Math.round(vc.lat * 10) / 10,
          lon: Math.round(vc.lon * 10) / 10,
          limbDeg: Math.round(limbDeg * 10) / 10
        };
        if (rvc.lat !== viewCenterRef.current.lat || rvc.lon !== viewCenterRef.current.lon ||
            rvc.limbDeg !== viewCenterRef.current.limbDeg) {
          viewCenterRef.current = rvc;
          setViewCenter(rvc);
        }
      }

      controls.update();

      // ズームに応じて月面写真の解像度を上げる（初期表示は2Kのまま、寄ったときだけ4K/8Kを読みに行く）
      const camDist = camera.position.distanceTo(controls.target);
      const wantTier = pickMoonTexTier(camDist);
      if (wantTier !== moonTexTierRef.current) {
        moonTexTierRef.current = wantTier;
        loadTexture(MOON_TEX_URL[wantTier]).then((tex) => {
          if (moonTexTierRef.current !== wantTier) return; // その間にさらにズームが変わっていたら古い結果は捨てる
          tex.colorSpace = THREE.SRGBColorSpace;
          tex.anisotropy = renderer.capabilities.getMaxAnisotropy();
          moonPhotoTextureRef.current = tex;
          const mat = moonMeshRef.current?.material;
          if (mat instanceof THREE.MeshStandardMaterial) {
            mat.bumpMap = tex;
            mat.bumpScale = 0.012;
            applyMoonTextureVisibility();
          }
        });
      }

      renderer.render(scene, camera);
    };
    animate();

    const handleResize = () => {
      if (!containerRef.current) return;
      const w = containerRef.current.clientWidth;
      const h = containerRef.current.clientHeight;
      camera.aspect = w / h;
      camera.updateProjectionMatrix();
      renderer.setSize(w, h);
    };
    const resizeObserver = new ResizeObserver(handleResize);
    resizeObserver.observe(containerRef.current);

    return () => {
      if (animationFrameIdRef.current) cancelAnimationFrame(animationFrameIdRef.current);
      resizeObserver.disconnect();
      controls.dispose();
      renderer.dispose();
      colorTexture.dispose();
      bumpTexture.dispose();
      photoTexture?.dispose();
      overlayMaterial.dispose();
      graticuleMat.dispose();
      earthTexture?.dispose();
      earthMesh.geometry.dispose();
      earthMesh.material.dispose();
      sunTexture?.dispose();
      sunMesh.geometry.dispose();
      sunMesh.material.dispose();
      sunGlowTexture.dispose();
      sunGlow.material.dispose();
      pinGroup.children.forEach((c) => {
        if (c instanceof THREE.Mesh) {
          c.geometry.dispose();
          (c.material as THREE.Material).dispose();
        }
      });
      textureCacheRef.current.forEach((p) => { p.then((tex) => tex.dispose()).catch(() => {}); });
      textureCacheRef.current.clear();
      diurnalTexturesRef.current?.forEach((tex) => tex.dispose());
      diurnalTexturesRef.current = null;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ワイヤーフレーム・明るさ・データ層表示・月面写真の表示。太陽の向きは世界座標で固定（月のほうが自転する）
  useEffect(() => {
    if (moonMeshRef.current?.material instanceof THREE.MeshStandardMaterial) {
      moonMeshRef.current.material.wireframe = settings.wireframe;
      moonMeshRef.current.material.needsUpdate = true;
    }
    if (sunLightRef.current) sunLightRef.current.intensity = settings.lightIntensity;
    if (overlayMeshRef.current) overlayMeshRef.current.visible = settings.showDataLayer;
    if (overlayGridGroupRef.current) overlayGridGroupRef.current.visible = settings.showDataLayer;
    applyMoonTextureVisibility();
  }, [settings.wireframe, settings.lightIntensity, settings.showDataLayer, settings.showMoonTexture]);

  // 自転角スライダー（settings.moonRotationDeg）を月本体の実際の回転に反映する。
  // 自動回転中は animate 側がこの値を進めつつ定期的に書き戻しているので、ここでの代入は
  // 「現在地に上書きする」だけで実質無害（手動でスライダーを動かしたときに効くのが主目的）。
  useEffect(() => {
    if (moonSpinGroupRef.current) {
      moonSpinGroupRef.current.rotation.y = THREE.MathUtils.degToRad(settings.moonRotationDeg);
    }
  }, [settings.moonRotationDeg]);

  // データ層の切り替え：静的な指標（site_environment 由来）か、1日のアニメーション（diviner 由来）か
  useEffect(() => {
    const material = overlayMaterialRef.current;
    if (!material) return;
    let cancelled = false;

    if (settings.dataLayerKey === DIURNAL_KEY) {
      setDiurnalLoading(true);
      // ホバー表示用の値ルックアップも同時に読みに行く（PNG24枚と同じく、この層を選んだときだけ）
      loadDiurnalLookup().then((data) => { diurnalLookupRef.current = data; });
      Promise.all(DIURNAL.frames.map((f) => loadTexture(f.texture)))
        .then((textures) => {
          if (cancelled) return;
          diurnalTexturesRef.current = textures;
          setDiurnalLoading(false);
          // 今の自転角（moonSpinGroup の実際の回転。ロード中に動いていてもズレないよう live 値を使う）
          const rotationDeg = moonSpinGroupRef.current
            ? THREE.MathUtils.radToDeg(moonSpinGroupRef.current.rotation.y)
            : settingsRef.current.moonRotationDeg;
          const idx = nearestDiurnalFrameIndexForRotation(rotationDeg);
          material.map = textures[idx];
          material.needsUpdate = true;
        })
        .catch(() => { if (!cancelled) setDiurnalLoading(false); });
    } else {
      const layer = DATA_LAYERS.find((l) => l.key === settings.dataLayerKey) ?? DATA_LAYERS[0];
      loadTexture(layer.texture).then((tex) => {
        if (cancelled) return;
        material.map = tex;
        material.needsUpdate = true;
      });
    }
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [settings.dataLayerKey]);

  // 1日のアニメーション表示中のフレーム切り替え自体は、animate ループが毎フレーム
  // moonSpinGroup の実際の回転角から直接おこなう（手動スライダー・自動回転のどちらでも
  // ズレない。上の useEffect は「初回ロード時の初期フレーム」だけを担当）。

  // 選択地点が変わったら、その方向にカメラを寄せる目標をセット
  useEffect(() => {
    if (!selectedFeature) return;
    const dir = latLongToVector3(selectedFeature.latitude, selectedFeature.longitude, 1).normalize();
    const sph = new THREE.Spherical().setFromVector3(dir);
    focusRef.current = { az: sph.theta, pol: sph.phi, dist: FOCUS_DISTANCE };
  }, [selectedFeature]);

  // 左下の展開図：カメラに今見えている範囲を淡く塗る（「シアンの点1つしか見えていないわけでは
  // ない」という指摘への対応）。1点だけでなく、実際に見えている角度（limbDeg）の範囲全体を
  // ピクセル単位の角距離判定で塗るので、球面上の本当の見え方（ズームすると狭くなる）と一致する。
  useEffect(() => {
    const canvas = miniMapCanvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const w = canvas.width;
    const h = canvas.height;
    const img = ctx.createImageData(w, h);
    const lat0 = (viewCenter.lat * Math.PI) / 180;
    const lon0 = (viewCenter.lon * Math.PI) / 180;
    const sinLat0 = Math.sin(lat0);
    const cosLat0 = Math.cos(lat0);
    const limbRad = (viewCenter.limbDeg * Math.PI) / 180;
    const cosLimb = Math.cos(limbRad);
    // 縁を少しだけぼかす（角距離で約4°ぶんフェード）。見えるか見えないかの硬い境界線より
    // 「だんだん端が見えにくくなる」実際の見え方に近い。
    const featherRad = (4 * Math.PI) / 180;
    const cosFeatherInner = Math.cos(Math.max(0, limbRad - featherRad));

    for (let py = 0; py < h; py++) {
      const lat = (Math.PI / 2) - ((py + 0.5) / h) * Math.PI;
      const sinLat = Math.sin(lat);
      const cosLat = Math.cos(lat);
      for (let px = 0; px < w; px++) {
        const lon = ((px + 0.5) / w) * 2 * Math.PI - Math.PI;
        const cosD = sinLat * sinLat0 + cosLat * cosLat0 * Math.cos(lon - lon0);
        let alpha = 0;
        if (cosD > cosLimb) {
          alpha = cosD >= cosFeatherInner ? 0.38
            : 0.38 * (cosD - cosLimb) / (cosFeatherInner - cosLimb);
        }
        const idx = (py * w + px) * 4;
        img.data[idx] = 34;      // cyan-400 相当
        img.data[idx + 1] = 211;
        img.data[idx + 2] = 238;
        img.data[idx + 3] = Math.round(alpha * 255);
      }
    }
    ctx.putImageData(img, 0, 0);
  }, [viewCenter]);

  // --- ポインタ操作。回転・ズームは OrbitControls。ここでは hover とクリック選択・ピン設置だけ ---
  const handlePointerMove = (e: React.PointerEvent<HTMLCanvasElement>) => {
    lastPointerClientRef.current = { x: e.clientX, y: e.clientY };
    const info = getHoverInfoAtClient(e.clientX, e.clientY);
    setHoverInfo(info);
    setTooltipPos(info ? { x: e.clientX, y: e.clientY } : null);
  };

  const handlePointerDown = (e: React.PointerEvent<HTMLCanvasElement>) => {
    pointerDownRef.current = { x: e.clientX, y: e.clientY, t: performance.now() };
  };

  const updatePinnedInfo = (info: HoverInfo | null) => {
    pinnedInfoRef.current = info;
    setPinnedInfo(info);
  };

  /** クリックした地点にピンを立てる（moonSpinGroup の子として置くので自転に追従する）。 */
  const placePinAt = (lat: number, lon: number) => {
    const pin = pinMarkerRef.current;
    if (!pin) return;
    const normal = latLongToVector3(lat, lon, 1).normalize();
    pin.position.copy(normal).multiplyScalar(MOON_RADIUS);
    pin.quaternion.setFromUnitVectors(new THREE.Vector3(0, 1, 0), normal);
    pin.visible = true;
  };

  const clearPin = () => {
    if (pinMarkerRef.current) pinMarkerRef.current.visible = false;
    updatePinnedInfo(null);
    setFollowPin(false); // ピンが無くなったら追従も解除
  };

  const handlePointerUp = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const down = pointerDownRef.current;
    pointerDownRef.current = null;
    if (!down) return;
    // ドラッグ（回転）とクリック（選択・ピン設置）を区別
    const moved = Math.hypot(e.clientX - down.x, e.clientY - down.y);
    if (moved > 5 || performance.now() - down.t > 500) return;
    const info = getHoverInfoAtClient(e.clientX, e.clientY);
    if (!info) return;
    // クリックしてある地点にピンを立て、そこのデータを表示する（要望への対応）。
    // 既知の地点の近くなら、従来どおりその地点として選択しカメラも寄せる。
    updatePinnedInfo(info);
    placePinAt(info.lat, info.lon);
    if (info.feature) onSelectFeature(info.feature);
  };

  const zoomBy = (factor: number) => {
    const c = controlsRef.current;
    const cam = cameraRef.current;
    if (!c || !cam) return;
    const offset = cam.position.clone().sub(c.target);
    const len = THREE.MathUtils.clamp(offset.length() * factor, MIN_DISTANCE, MAX_DISTANCE);
    cam.position.copy(c.target).add(offset.setLength(len));
    c.update();
  };

  const resetView = () => {
    focusRef.current = null;
    controlsRef.current?.reset();
  };

  // データ層の凡例に出す値（静的な層／1日のアニメーションのどちらを選んでいるかで出し分け）
  const isDiurnal = settings.dataLayerKey === DIURNAL_KEY;
  const currentLayer = DATA_LAYERS.find((l) => l.key === settings.dataLayerKey);
  const legendKey = isDiurnal ? DIURNAL_KEY : currentLayer?.key ?? '';
  const legendRawUnit = isDiurnal ? DIURNAL.unit : currentLayer?.unit ?? '';
  const legendMinDisp = toCelsiusDisplay(legendKey, legendRawUnit, isDiurnal ? DIURNAL.min : currentLayer?.min ?? 0);
  const legendMaxDisp = toCelsiusDisplay(legendKey, legendRawUnit, isDiurnal ? DIURNAL.max : currentLayer?.max ?? 0);
  const legendLabel = isDiurnal ? DIURNAL.label : currentLayer?.label ?? '';
  const legendUnit = legendMinDisp.unit;
  const legendMin = legendUnit === '℃' ? Math.round(legendMinDisp.value) : legendMinDisp.value;
  const legendMax = legendUnit === '℃' ? Math.round(legendMaxDisp.value) : legendMaxDisp.value;
  const legendDesc = isDiurnal ? DIURNAL.desc : currentLayer?.desc ?? '';
  const legendSource = isDiurnal ? DIURNAL.source : currentLayer?.source ?? '';
  const legendGradient = isDiurnal ? DIURNAL.gradientCss : currentLayer?.gradientCss ?? '';

  return (
    <div
      ref={containerRef}
      id="moon-3d-viewport-container"
      className={
        fill
          ? 'relative w-full h-full min-h-[320px] bg-slate-950 overflow-hidden flex flex-col select-none'
          : 'relative w-full h-full min-h-[460px] bg-slate-950 rounded-2xl overflow-hidden border border-slate-800 shadow-2xl flex flex-col select-none'
      }
    >
      <canvas
        ref={canvasRef}
        id="moon-threejs-canvas"
        onPointerDown={handlePointerDown}
        onPointerMove={handlePointerMove}
        onPointerUp={handlePointerUp}
        onPointerLeave={() => { lastPointerClientRef.current = null; setHoverInfo(null); setTooltipPos(null); }}
        className="w-full h-full cursor-grab active:cursor-grabbing block touch-none"
      />

      {/* HUD header */}
      <div className="absolute top-4 left-4 right-4 flex items-center justify-between pointer-events-none">
        <div className="flex items-center gap-2 bg-slate-900/80 backdrop-blur-md border border-slate-700/60 px-3 py-1.5 rounded-xl shadow-lg pointer-events-auto">
          <div className="w-2.5 h-2.5 rounded-full bg-cyan-400 animate-pulse" />
          <span className="text-xs font-semibold tracking-wide text-slate-200">3D 月球儀</span>
          <span className="text-[11px] text-slate-400 pl-1 border-l border-slate-700">ドラッグで回す</span>
        </div>

        <div className="hidden md:flex items-center gap-3 bg-slate-900/80 backdrop-blur-md border border-slate-700/60 px-3.5 py-1.5 rounded-xl shadow-lg text-xs font-mono text-slate-300 pointer-events-auto">
          <Compass className="w-4 h-4 text-cyan-400" />
          {hoverInfo ? (
            <span>
              Lat: <strong className="text-white">{hoverInfo.lat > 0 ? `+${hoverInfo.lat}` : hoverInfo.lat}°</strong> |
              Lon: <strong className="text-white">{hoverInfo.lon > 0 ? `+${hoverInfo.lon}` : hoverInfo.lon}°</strong>
            </span>
          ) : selectedFeature ? (
            <span>
              選択中: <strong className="text-white">{selectedFeature.nameJa}</strong>（{selectedFeature.latitude}°, {selectedFeature.longitude}°）
            </span>
          ) : (
            <span className="text-slate-400">ドラッグで回転 ・ ホイールで拡大縮小</span>
          )}
        </div>
      </div>

      {/* データ層の凡例＋切り替え。requirements_v3.3 続き（「地球の風」Phase1〜3） */}
      {settings.showDataLayer && (
        <div
          id="data-layer-legend"
          className="absolute top-16 left-4 bg-slate-900/85 backdrop-blur-md border border-rose-500/30 px-3.5 py-2.5 rounded-2xl shadow-xl flex flex-col gap-1.5 text-xs text-slate-300 pointer-events-auto max-w-[230px]"
        >
          <div className="flex items-center gap-1.5 text-rose-300 font-semibold">
            <Thermometer className="w-3.5 h-3.5 shrink-0" />
            <select
              id="select-data-layer"
              value={settings.dataLayerKey}
              onChange={(e) => onUpdateSettings({ dataLayerKey: e.target.value })}
              className="bg-slate-800 border border-rose-500/30 rounded-lg px-1.5 py-0.5 text-[11px] text-rose-200 font-semibold flex-1 min-w-0"
            >
              {DATA_LAYERS.map((l) => (
                <option key={l.key} value={l.key}>{l.label}</option>
              ))}
              <option value={DIURNAL_KEY}>{DIURNAL.label}</option>
            </select>
          </div>
          <div className="h-2.5 rounded-full" style={{ background: legendGradient }} />
          <div className="flex justify-between text-[10px] font-mono text-slate-400">
            <span>{legendMin}{legendUnit}</span>
            <span>{legendMax}{legendUnit}</span>
          </div>
          <p className="text-[10px] text-slate-400 leading-snug">{legendDesc}</p>
          {isDiurnal && (
            <p className="text-[10px] text-cyan-300 leading-snug">
              {diurnalLoading
                ? `読み込み中…（${DIURNAL.nFrames}枚）`
                : `◀ 左下の「月の自転」スライダーか、自動回転（▶）で時間が進みます`}
            </p>
          )}
          <p className="text-[10px] text-slate-500">出典：{legendSource}</p>
        </div>
      )}

      {/* ピン留めした地点（クリックで設置。マウスを離しても消えない）。要望への対応 */}
      {pinnedInfo && (
        <div
          id="pinned-point-card"
          className="absolute top-16 right-4 bg-slate-900/90 backdrop-blur-md border border-amber-500/40 px-3.5 py-2.5 rounded-2xl shadow-xl flex flex-col gap-1 text-xs text-slate-300 pointer-events-auto max-w-[230px]"
        >
          <div className="flex items-center justify-between gap-2">
            <span className="text-amber-300 font-semibold text-[11px]">📌 ピン留めした地点</span>
            <div className="flex items-center gap-1">
              <button
                id="btn-toggle-follow-pin"
                onClick={() => setFollowPin((v) => !v)}
                title={followPin ? '自転してもピンを追いかけるのを止める' : '自転してもピンを見失わないようにする'}
                className={`flex items-center gap-0.5 px-1.5 py-0.5 rounded-md border text-[10px] transition-colors ${
                  followPin
                    ? 'bg-cyan-500/20 text-cyan-300 border-cyan-500/40'
                    : 'text-slate-400 border-slate-600 hover:text-slate-200'
                }`}
              >
                <Crosshair className="w-3 h-3" />
                追従
              </button>
              <button
                id="btn-clear-pin"
                onClick={clearPin}
                title="ピンを消す"
                className="text-slate-500 hover:text-slate-200 leading-none px-1"
              >
                ✕
              </button>
            </div>
          </div>
          {pinnedInfo.feature && (
            <div className="text-sm font-semibold text-slate-100">
              {CATEGORY_EMOJI[pinnedInfo.feature.category] ?? '📍'} {pinnedInfo.feature.nameJa}
            </div>
          )}
          <div className="text-xs text-slate-400 font-mono">
            Lat: {pinnedInfo.lat}° | Lon: {pinnedInfo.lon}°
          </div>
          {pinnedInfo.layerValue && (
            <div className="text-xs text-rose-300 font-mono">
              {pinnedInfo.layerValue.label}: <strong>{pinnedInfo.layerValue.value.toFixed(1)}{pinnedInfo.layerValue.unit}</strong>
            </div>
          )}
        </div>
      )}

      {/* Toolbar */}
      <div className="absolute bottom-4 right-4 flex flex-col gap-2 bg-slate-900/85 backdrop-blur-md border border-slate-700/70 p-2 rounded-2xl shadow-xl pointer-events-auto">
        <button
          id="btn-toggle-autorotate"
          onClick={() => onUpdateSettings({ autoRotate: !settings.autoRotate })}
          title={settings.autoRotate ? '月の自転を止める' : '月を自転させる（地球も一緒に空を巡ります）'}
          className={`p-2.5 rounded-xl text-xs flex items-center justify-center transition-all ${
            settings.autoRotate ? 'bg-cyan-500/20 text-cyan-300 border border-cyan-500/40' : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800'
          }`}
        >
          {settings.autoRotate ? <Pause className="w-4 h-4" /> : <Play className="w-4 h-4" />}
        </button>

        <button
          id="btn-toggle-moon-texture"
          onClick={() => onUpdateSettings({ showMoonTexture: !settings.showMoonTexture })}
          title={settings.showMoonTexture ? '月の実写画像を消す（データ層の色と紛れないように）' : '月の実写画像を表示する'}
          className={`p-2.5 rounded-xl text-xs flex items-center justify-center transition-all ${
            settings.showMoonTexture ? 'bg-amber-500/20 text-amber-300 border border-amber-500/40' : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800'
          }`}
        >
          <Image className="w-4 h-4" />
        </button>

        <button
          id="btn-toggle-data-layer"
          onClick={() => onUpdateSettings({ showDataLayer: !settings.showDataLayer })}
          title={`データ層：${legendLabel}を重ねる`}
          className={`p-2.5 rounded-xl text-xs flex items-center justify-center transition-all ${
            settings.showDataLayer ? 'bg-rose-500/20 text-rose-300 border border-rose-500/40' : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800'
          }`}
        >
          <Thermometer className="w-4 h-4" />
        </button>

        <button
          id="btn-toggle-wireframe"
          onClick={() => onUpdateSettings({ wireframe: !settings.wireframe })}
          title="ワイヤーフレーム"
          className={`p-2.5 rounded-xl text-xs flex items-center justify-center transition-all ${
            settings.wireframe ? 'bg-indigo-500/20 text-indigo-300 border border-indigo-500/40' : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800'
          }`}
        >
          <Maximize2 className="w-4 h-4" />
        </button>

        <div className="h-px bg-slate-700/80 my-0.5" />

        <button id="btn-zoom-in" onClick={() => zoomBy(0.8)} title="拡大"
          className="p-2.5 rounded-xl text-slate-400 hover:text-slate-200 hover:bg-slate-800 transition-all">
          <ZoomIn className="w-4 h-4" />
        </button>
        <button id="btn-zoom-out" onClick={() => zoomBy(1.25)} title="縮小"
          className="p-2.5 rounded-xl text-slate-400 hover:text-slate-200 hover:bg-slate-800 transition-all">
          <ZoomOut className="w-4 h-4" />
        </button>
        <button id="btn-reset-view" onClick={resetView} title="視点を初期位置に戻す"
          className="p-2.5 rounded-xl text-slate-400 hover:text-cyan-300 hover:bg-slate-800 transition-all">
          <RotateCcw className="w-4 h-4" />
        </button>
      </div>

      {/* 左下：展開図（今の向き・ピンの場所）＋自転スライダー。要望「月球儀が今どこを向いているか、
          ピンがどこに刺さったか見られるように」への対応 */}
      <div className="absolute bottom-4 left-4 flex flex-col gap-2 pointer-events-none">
        <div
          id="mini-map-card"
          className="bg-slate-900/85 backdrop-blur-md border border-slate-700/70 p-2 rounded-2xl shadow-xl pointer-events-auto flex flex-col gap-1.5"
        >
          <div className="flex items-center gap-1.5 text-[10px] text-slate-400 px-0.5">
            <Compass className="w-3 h-3 text-cyan-400 shrink-0" />
            <span>展開図（水色＝今見えている範囲{pinnedInfo ? '・●ピン' : ''}）</span>
          </div>
          <div
            id="mini-map"
            className="relative w-40 sm:w-48 rounded-lg overflow-hidden border border-slate-700/60"
            style={{ aspectRatio: '2 / 1' }}
          >
            {/* 月面テクスチャと同じ正距円筒図法（x=0が経度-180°、y=0が緯度+90°）。
                いわゆる「メルカトル図法」だと極が無限に伸びて使いにくいので、この投影法にしている。 */}
            <img
              src="textures/moon_lroc_color_2k.jpg"
              alt="月面の展開図（正距円筒図法）"
              className="absolute inset-0 w-full h-full object-cover"
              draggable={false}
            />
            {/* 「点1つしか見えていないわけではない」という指摘への対応：カメラの視野に実際に
                入っている範囲全体を淡く塗る（ズームすると狭く、離れると広くなる）。 */}
            <canvas
              ref={miniMapCanvasRef}
              width={200}
              height={100}
              className="absolute inset-0 w-full h-full pointer-events-none"
            />
            <div
              id="mini-map-view-marker"
              className="absolute w-2 h-2 -translate-x-1/2 -translate-y-1/2 rounded-full border-2 border-cyan-300"
              style={{
                left: `${((viewCenter.lon + 180) / 360) * 100}%`,
                top: `${((90 - viewCenter.lat) / 180) * 100}%`,
                boxShadow: '0 0 4px rgba(34,211,238,0.9)'
              }}
              title="今カメラが真正面から見ている地点（中心）"
            />
            {pinnedInfo && (
              <div
                id="mini-map-pin-marker"
                className="absolute w-2.5 h-2.5 -translate-x-1/2 -translate-y-1/2 rounded-full bg-amber-400 border border-amber-100"
                style={{
                  left: `${((pinnedInfo.lon + 180) / 360) * 100}%`,
                  top: `${((90 - pinnedInfo.lat) / 180) * 100}%`,
                  boxShadow: '0 0 4px rgba(250,204,21,0.9)'
                }}
                title="ピン留めした地点"
              />
            )}
          </div>
        </div>

        {/* 自転（＝1日の時刻）スライダー＋自転の速さ。太陽の向きは固定で、月を回すことで昼夜が移り変わる */}
        <div className="bg-slate-900/85 backdrop-blur-md border border-slate-700/70 px-3.5 py-2.5 rounded-2xl shadow-xl flex items-center gap-3 text-xs text-slate-300 pointer-events-auto">
          <RotateCw className="w-4 h-4 text-amber-400 shrink-0" />
          <div className="flex flex-col gap-2">
            <div className="flex flex-col gap-1">
              <div className="flex justify-between items-center text-[10px] text-slate-400">
                <span>月の自転（1日の時刻）</span>
                <span className="font-mono text-amber-300">{Math.round(settings.moonRotationDeg)}°</span>
              </div>
              <input
                id="slider-moon-rotation"
                type="range" min="0" max="360" step="5"
                value={settings.moonRotationDeg}
                onChange={(e) => onUpdateSettings({ moonRotationDeg: Number(e.target.value) })}
                className="w-28 sm:w-36 h-1.5 bg-slate-700 rounded-lg appearance-none cursor-pointer accent-amber-400"
              />
            </div>
            <div className="flex flex-col gap-1">
              <div className="flex justify-between items-center text-[10px] text-slate-400">
                <span>自転の速さ（自動回転中）</span>
                <span className="font-mono text-amber-300">×{settings.rotationSpeed.toFixed(1)}</span>
              </div>
              <input
                id="slider-rotation-speed"
                type="range" min="0.1" max="3" step="0.1"
                value={settings.rotationSpeed}
                onChange={(e) => onUpdateSettings({ rotationSpeed: Number(e.target.value) })}
                className="w-28 sm:w-36 h-1.5 bg-slate-700 rounded-lg appearance-none cursor-pointer accent-amber-400"
              />
            </div>
          </div>
        </div>
      </div>

      {/* Hover tooltip：常設ピンの代わりに、カーソルを合わせた地点の情報をまとめて出す */}
      {hoverInfo && tooltipPos && (
        <div
          id="moon-hover-tooltip"
          style={{ left: `${tooltipPos.x + 14}px`, top: `${tooltipPos.y + 14}px`, position: 'fixed' }}
          className="z-50 pointer-events-none bg-slate-900/95 backdrop-blur-lg border border-cyan-500/40 p-3 rounded-xl shadow-2xl max-w-xs animate-in fade-in"
        >
          {hoverInfo.feature && (
            <>
              <div className="text-[10px] uppercase font-bold tracking-wider text-cyan-400 mb-0.5">
                {CATEGORY_EMOJI[hoverInfo.feature.category] ?? '📍'}
              </div>
              <div className="text-sm font-semibold text-slate-100">{hoverInfo.feature.nameJa}</div>
            </>
          )}
          <div className="text-xs text-slate-400 font-mono mt-1">
            Lat: {hoverInfo.lat}° | Lon: {hoverInfo.lon}°
          </div>
          {hoverInfo.layerValue && (
            <div className="text-xs text-rose-300 font-mono mt-1">
              {hoverInfo.layerValue.label}: <strong>{hoverInfo.layerValue.value.toFixed(1)}{hoverInfo.layerValue.unit}</strong>
            </div>
          )}
          {hoverInfo.feature && (
            <div className="text-[10px] text-cyan-300 font-medium mt-2">クリックで選択 →</div>
          )}
        </div>
      )}
    </div>
  );
};
