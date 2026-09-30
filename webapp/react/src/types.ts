export type LunarDatasetCategory = 'missions' | 'geomorphology' | 'geophysics' | 'resources';

export interface LunarFeature {
  id: string;
  name: string;
  nameJa: string;
  category: string;
  latitude: number;
  longitude: number;
  diameterKm?: number;
  depthKm?: number;
  year?: string | number;
  mission?: string;
  agency?: string;
  status?: string;
  description: string;
  attributes: Record<string, string | number | boolean>;
}

export interface DatasetColumn {
  key: string;
  label: string;
  unit?: string;
}

export interface LunarDataset {
  id: string;
  title: string;
  titleJa: string;
  description: string;
  category: LunarDatasetCategory;
  categoryLabelJa: string;
  icon: string;
  badgeColor: string;
  columns: DatasetColumn[];
  data: LunarFeature[];
  downloadFileName: string;
}

export interface MoonViewerSettings {
  autoRotate: boolean;
  rotationSpeed: number;
  wireframe: boolean;
  lightIntensity: number;
  /** 月の自転角（0〜360度）。太陽は世界座標で固定し、この角度だけ月本体を回すことで昼夜が
   *  移り変わる（＝「1日の温度アニメーション」のどのフレームを見せるかも、この角度から決まる）。
   *  旧「太陽光照射角」を統合したもの（自転と昼夜を別々の概念にしないための単純化）。 */
  moonRotationDeg: number;
  selectedCategoryFilter: string | 'all';
  /** 月面にデータ層（例：1日の温度差）を半透明で重ねるか（requirements_v3.3 続き・地球の風 Phase1） */
  showDataLayer: boolean;
  /** 表示中のデータ層のキー（site_environment.csv の列名、または 'diurnal_temp'＝1日のアニメーション） */
  dataLayerKey: string;
  /** 月の実写テクスチャ（色）を表示するか。オフにすると素の灰色の球になり、データ層の色が
   *  月面写真自体の色と紛れない（「月の色なのかデータなのか分からない」という指摘への対応）。 */
  showMoonTexture: boolean;
}
