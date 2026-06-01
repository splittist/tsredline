export { assignUnids } from './assignUnids';
export { compareDocx } from './compareDocx';
export { createAtomList } from './createAtomList';
export { getComparisonUnitList } from './getComparisonUnitList';
export { hashBlockLevelContent } from './hashBlockLevelContent';
export { preprocessDocx } from './preprocessDocx';
export type {
  CompareOptions,
  CompareResult,
  ComparisonChange,
  ComparisonMetadata,
  ComparisonNotice,
} from './types';
export type { AssignUnidsResult } from './assignUnids';
export type {
  AtomKind,
  AtomListResult,
  ComparisonUnitAtom,
  CreateAtomListOptions,
} from './createAtomList';
export type {
  BlockHash,
  HashBlockLevelOptions,
  HashBlockLevelResult,
} from './hashBlockLevelContent';
export type {
  ComparisonUnit,
  ComparisonUnitGroup,
  ComparisonUnitGroupKind,
  ComparisonUnitWord,
} from './getComparisonUnitList';
export type { PreprocessedDocx } from './preprocessDocx';
export type {
  PreprocessDocxOptions,
  RevisionMode,
} from './preprocessDocx';
