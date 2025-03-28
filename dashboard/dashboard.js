// Google Apps script web/app 배포시 사용되는 자바스크립트(실제론 gs확장자)

/**
 * 배차 대시보드 앱 - Google Apps Script로 구현한 배차 실적 시각화 대시보드
 * 실적 탭에 있는 데이터를 시각화하여 웹앱으로 제공합니다.
 */

// 웹앱의 기본 HTML을 생성하는 함수
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('배차 대시보드')
    .setFaviconUrl('https://ssl.gstatic.com/docs/script/images/favicon.ico')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// HTML 파일에 포함시킬 JavaScript 및 CSS 파일을 가져오는 함수
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 실적 데이터를 가져오는 함수
function getDispatchData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('실적'); // 실적 시트명을 적절히 변경하세요
  
  // 데이터 범위 (헤더 포함)
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // 헤더 행 제외
  const headers = values[0];
  
  // 빈 행 제거 및 유효한 데이터만 필터링
  const data = values.slice(1).filter(row => {
    // 최소한 담당자 필드가 있는 행만 유효한 데이터로 간주
    return row[2] && row[2].toString().trim() !== '';
  });
  
  // 실제 데이터 건수 로그
  console.log('유효한 데이터 건수: ' + data.length);
  
  return {
    headers: headers,
    data: data
  };
}

// 배차 담당자별 실적 통계를 계산하는 함수
function getDispatcherStats() {
  const { data } = getDispatchData();
  
  // 담당자 인덱스 (열 번호)
  const dispatcherIdx = 2;  // '담당자' 열 (0부터 시작하므로 3번째 열은 2)
  const resultIdx = 5;      // '배차 결과' 열
  const timeSpentIdx = 20;  // '소요시간' 열 (U열) -> 0부터 시작하므로 21번째 열은 20
  
  // 점수 절반으로 계산할 특정 담당자 리스트
  const halfScoreDispatchers = [
    // 여기에 점수를 절반으로 계산할 담당자 이름을 추가하세요
    // 예: "홍길동", "김철수", "이영희"
  ];
  
  // 담당자별 통계
  const stats = {};
  
  // 근무자 실적 데이터 가져오기
  const workPerformanceData = getWorkPerformanceCachedData();
  
  data.forEach(row => {
    const dispatcher = row[dispatcherIdx];
    const result = row[resultIdx];
    
    // 소요시간 처리
    let timeSpent = row[timeSpentIdx];
    let timeSeconds = 0;
    
    // 날짜 객체인 경우 시, 분, 초 추출
    if (timeSpent instanceof Date) {
      timeSeconds = timeSpent.getHours() * 3600 + timeSpent.getMinutes() * 60 + timeSpent.getSeconds();
    } 
    // 문자열인 경우 H:mm:SS 또는 오전/오후 형식 파싱
    else if (typeof timeSpent === 'string') {
      timeSeconds = parseTimeString(timeSpent);
    }
    // 숫자인 경우 그대로 사용
    else if (typeof timeSpent === 'number' && !isNaN(timeSpent)) {
      timeSeconds = timeSpent;
    }
    
    // 담당자가 통계에 없으면 초기화
    if (!stats[dispatcher]) {
      stats[dispatcher] = {
        total: 0,
        totalScore: 0,
        systemDispatch: 0,
        operatorDispatch: 0,
        normalMoving: 0,
        cantProcess: 0,
        other: 0,
        totalTime: 0,
        avgTime: 0,
        timeSamples: [], // 개별 소요시간 샘플 저장
        isHalfScore: halfScoreDispatchers.includes(dispatcher), // 특정 리스트에 포함되는지 여부
        workShift: '', // 근무 시프트
        workHours: '', // 근무 시간
        goalAchievement: '' // 목표 달성
      };
      
      // 근무자 실적 데이터에서 해당 담당자 정보 찾기
      if (workPerformanceData && workPerformanceData[dispatcher.toString().toLowerCase().trim()]) {
        stats[dispatcher].workShift = workPerformanceData[dispatcher.toString().toLowerCase().trim()].workShift;
        stats[dispatcher].workHours = workPerformanceData[dispatcher.toString().toLowerCase().trim()].workHours;
        stats[dispatcher].goalAchievement = workPerformanceData[dispatcher.toString().toLowerCase().trim()].goalAchievement;
      }
    }
    
    // 총 건수 증가
    stats[dispatcher].total++;
    
    // 점수 계산 (오퍼 배차 1점, 나머지 0.5점, 기타 0점)
    let score = 0;
    
    // 배차 결과에 따른 분류 및 점수 계산
    if (result.includes('시스템 배차')) {
      stats[dispatcher].systemDispatch++;
      score = 0.5;
    } else if (result.includes('오퍼레이터 배차')) {
      stats[dispatcher].operatorDispatch++;
      score = 1;
    } else if (result.includes('정상 이동 중으로 확인')) {
      stats[dispatcher].normalMoving++;
      score = 0.5;
    } else if (result.includes('처리 불가')) {
      stats[dispatcher].cantProcess++;
      score = 0.5;
    } else {
      stats[dispatcher].other++;
      score = 0;
    }
    
    // 특정 리스트에 포함된 담당자는 점수를 절반으로 계산
    if (stats[dispatcher].isHalfScore) {
      score = score / 2;
    }
    
    // 총 점수 누적
    stats[dispatcher].totalScore += score;
    
    // 소요시간이 유효한 값이면 누적 및 샘플 저장
    if (timeSeconds > 0) {
      stats[dispatcher].totalTime += timeSeconds;
      stats[dispatcher].timeSamples.push(timeSeconds);
    }
  });
  
  // 평균 시간 계산
  Object.keys(stats).forEach(dispatcher => {
    const validTimes = stats[dispatcher].timeSamples.filter(time => time <= 1200); // 20분(1200초) 이하만 유효
    if (validTimes.length > 0) {
      const validTotal = validTimes.reduce((sum, time) => sum + time, 0);
      stats[dispatcher].avgTime = validTotal / validTimes.length;
    }
    
    // 점수 소수점 첫째 자리까지 반올림
    stats[dispatcher].totalScore = Math.round(stats[dispatcher].totalScore * 10) / 10;
  });
  
  return stats;
}

// 근무자 실적 데이터를 가져오는 함수
function getWorkPerformanceData() {
  try {
    // 근무자 실적 스프레드시트 URL (하드코딩)
    const workPerformanceUrl = "여기에 스프레드시트 URL을 입력하세요"; // 이 부분을 실제 URL로 변경하세요
    
    let ss;
    
    // URL이 있으면 해당 URL의 스프레드시트 열기, 없으면 현재 활성 스프레드시트 사용
    if (workPerformanceUrl && workPerformanceUrl !== "여기에 스프레드시트 URL을 입력하세요") {
      try {
        ss = SpreadsheetApp.openByUrl(workPerformanceUrl);
      } catch (e) {
        console.error('근무자 실적 URL로 스프레드시트를 열 수 없습니다:', e);
        ss = SpreadsheetApp.getActiveSpreadsheet();
      }
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    
    const sheet = ss.getSheetByName('근무자 실적'); // 근무자 실적 시트
    
    if (!sheet) {
      console.error('근무자 실적 시트를 찾을 수 없습니다.');
      return null;
    }
    
    // 데이터 범위 (헤더 포함)
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // 헤더 행 제외
    const data = values.slice(1);
    
    // 담당자별 데이터 매핑 (대소문자 구분 없이)
    const workPerformanceMap = {};
    
    data.forEach(row => {
      const dispatcher = row[1]; // B열: 상담사(담당자) 이름
      if (dispatcher && dispatcher.toString().trim() !== '') {
        // 소문자로 변환하여 저장 (대소문자 구분 없이 매칭하기 위함)
        const dispatcherLower = dispatcher.toString().toLowerCase().trim();
        workPerformanceMap[dispatcherLower] = {
          workShift: row[2],   // C열: 근무 시프트
          workHours: row[3],   // D열: 근무 시간
          goalAchievement: row[51]  // AZ열: 목표 달성
        };
      }
    });
    
    // 마지막 업데이트 시간 저장
    PropertiesService.getScriptProperties().setProperty('lastWorkPerformanceUpdate', new Date().toISOString());
    
    // 캐시 데이터 저장
    const props = PropertiesService.getScriptProperties();
    props.setProperty('workPerformanceCache', JSON.stringify(workPerformanceMap));
    
    return workPerformanceMap;
  } catch (error) {
    console.error('근무자 실적 데이터 로드 중 오류:', error);
    return null;
  }
}

// 근무자 실적 데이터 캐시를 관리하는 함수
function getWorkPerformanceCachedData() {
  // 스크립트 프로퍼티에서 마지막 업데이트 시간 가져오기
  const props = PropertiesService.getScriptProperties();
  const lastUpdateStr = props.getProperty('lastWorkPerformanceUpdate');
  
  // 캐시된 데이터가 없거나 4시간(14400000밀리초)이 지났으면 새로 가져오기
  if (!lastUpdateStr || isDataStale(lastUpdateStr, 14400000)) {
    console.log('근무자 실적 데이터 새로 가져오기');
    return getWorkPerformanceData();
  }
  
  // 캐시된 데이터 사용
  try {
    const cachedDataStr = props.getProperty('workPerformanceCache');
    if (!cachedDataStr) {
      return getWorkPerformanceData();
    }
    
    return JSON.parse(cachedDataStr);
  } catch (error) {
    console.error('캐시된 근무자 실적 데이터 파싱 오류:', error);
    return getWorkPerformanceData();
  }
}

// 데이터가 오래되었는지 확인하는 함수
function isDataStale(lastUpdateStr, maxAgeMs) {
  try {
    const lastUpdate = new Date(lastUpdateStr);
    const now = new Date();
    return (now.getTime() - lastUpdate.getTime()) > maxAgeMs;
  } catch (error) {
    console.error('날짜 비교 오류:', error);
    return true; // 오류 발생 시 데이터를 오래된 것으로 간주
  }
}

// 근무자 실적 데이터 강제 새로고침 함수
function forceRefreshWorkPerformanceData() {
  return getWorkPerformanceData();
}

// 시간 문자열을 초로 변환하는 함수
function parseTimeString(timeStr) {
  if (!timeStr || typeof timeStr !== 'string' || timeStr.trim() === '') return 0;
  
  const trimmedStr = timeStr.trim();
  
  // 한국어 시간 형식 처리 (예: "오전 12:03:29" 또는 "오후 3:15:20")
  const koreanTimeRegex = /^(오전|오후)\s+(\d+):(\d+):(\d+)$/;
  const koreanMatch = trimmedStr.match(koreanTimeRegex);
  
  if (koreanMatch) {
    let hours = parseInt(koreanMatch[2], 10);
    const minutes = parseInt(koreanMatch[3], 10);
    const seconds = parseInt(koreanMatch[4], 10);
    
    // 오후인 경우 12시간 추가 (단, 오후 12시는 그대로 12시)
    if (koreanMatch[1] === '오후' && hours < 12) {
      hours += 12;
    }
    // 오전 12시는 0시로 변환
    else if (koreanMatch[1] === '오전' && hours === 12) {
      hours = 0;
    }
    
    return hours * 3600 + minutes * 60 + seconds;
  }
  
  // H:mm:SS 형식 검사
  const timeRegex = /^(\d+):(\d{1,2}):(\d{1,2})$/;
  const match = trimmedStr.match(timeRegex);
  
  if (match) {
    const hours = parseInt(match[1], 10);
    const minutes = parseInt(match[2], 10);
    const seconds = parseInt(match[3], 10);
    
    return hours * 3600 + minutes * 60 + seconds;
  }
  
  // MM:SS 형식 검사
  const shortTimeRegex = /^(\d+):(\d{1,2})$/;
  const shortMatch = trimmedStr.match(shortTimeRegex);
  
  if (shortMatch) {
    const minutes = parseInt(shortMatch[1], 10);
    const seconds = parseInt(shortMatch[2], 10);
    
    return minutes * 60 + seconds;
  }
  
  // 숫자만 있는 경우 초로 간주
  if (!isNaN(trimmedStr)) {
    return parseInt(trimmedStr, 10);
  }
  
  return 0;
}

// 전체 배차 실적 요약 통계를 계산하는 함수
function getDispatchSummary() {
  const { data } = getDispatchData();
  
  // 필요한 열 인덱스
  const resultIdx = 5;     // '배차 결과' 열
  const delayTimeIdx = 4;  // '지연시간' 열
  const regionIdx = 13;    // 'region_id' 열
  const adTypeIdx = 15;    // 'AD 타입' 열
  
  // 요약 통계 객체
  const summary = {
    total: data.length,
    dispatchResults: {},
    regions: {},
    adTypes: {},
    avgDelayTime: 0,
    totalDelayTime: 0
  };
  
  // 데이터 집계
  data.forEach(row => {
    const result = row[resultIdx];
    const delayTime = row[delayTimeIdx];
    const region = row[regionIdx];
    const adType = row[adTypeIdx];
    
    // 배차 결과 집계
    if (!summary.dispatchResults[result]) {
      summary.dispatchResults[result] = 0;
    }
    summary.dispatchResults[result]++;
    
    // 지역 집계
    if (!summary.regions[region]) {
      summary.regions[region] = 0;
    }
    summary.regions[region]++;
    
    // AD 타입 집계
    if (!summary.adTypes[adType]) {
      summary.adTypes[adType] = 0;
    }
    summary.adTypes[adType]++;
    
    // 지연 시간 집계
    if (!isNaN(delayTime) && delayTime > 0) {
      summary.totalDelayTime += delayTime;
    }
  });
  
  // 평균 지연 시간 계산
  if (data.length > 0) {
    summary.avgDelayTime = summary.totalDelayTime / data.length;
  }
  
  return summary;
}

// 성과 점수 계산을 위한 가중치 설정 (전역 변수로 관리)
const PERFORMANCE_WEIGHTS = {
  totalScore: 0.2,    // 총 점수 가중치 (20%)
  operatorRatio: 0.5, // 오퍼레이터 배차 비중 가중치 (50%)
  avgTime: 0.3        // 평균 소요시간 가중치 (30%)
};

// 성과 점수 계산 공식 문자열 생성 함수
function getPerformanceFormulaText() {
  return `성과 점수 계산 공식: 총 점수(${PERFORMANCE_WEIGHTS.totalScore * 100}%) + 오퍼레이터 배차 비중(${PERFORMANCE_WEIGHTS.operatorRatio * 100}%) + 평균 소요시간 역수(${PERFORMANCE_WEIGHTS.avgTime * 100}%)`;
}

// 툴팁 업데이트 함수
function updatePerformanceTooltips() {
  const formulaText = getPerformanceFormulaText();
  const tooltipElements = document.querySelectorAll('th[data-tooltip="performance-formula"]');
  
  tooltipElements.forEach(element => {
    element.setAttribute('title', formulaText);
  });
}

// 고성과자/저성과자 통계를 계산하는 함수
function getTopPerformers() {
  const dispatcherStats = getDispatcherStats();
  
  // 담당자 데이터를 배열로 변환
  const dispatchersArray = Object.keys(dispatcherStats).map(dispatcher => {
    const stats = dispatcherStats[dispatcher];
    
    // 오퍼레이터 배차 비중 계산
    const operatorRatio = stats.total > 0 ? stats.operatorDispatch / stats.total : 0;
    
    return {
      dispatcher: dispatcher,
      total: stats.total,
      totalScore: stats.totalScore,
      systemDispatch: stats.systemDispatch,
      operatorDispatch: stats.operatorDispatch,
      operatorRatio: operatorRatio,
      normalMoving: stats.normalMoving,
      cantProcess: stats.cantProcess,
      other: stats.other,
      avgTime: stats.avgTime,
      // 원시 데이터 저장 (정규화 전)
      rawData: {
        totalScore: stats.totalScore,
        operatorRatio: operatorRatio,
        avgTime: stats.avgTime
      }
    };
  });
  
  // 최대값 찾기 (정규화를 위해)
  const maxTotalScore = Math.max(...dispatchersArray.map(item => item.rawData.totalScore));
  const maxOperatorRatio = 1.0; // 오퍼레이터 비중은 최대 100%
  const maxAvgTime = 300; // 최대 5분(300초)을 기준으로 함
  
  // 정규화 및 성과 점수 계산
  dispatchersArray.forEach(item => {
    // 각 요소 정규화 (0~1 사이 값으로 변환)
    const normalizedScore = maxTotalScore > 0 ? item.rawData.totalScore / maxTotalScore : 0;
    const normalizedRatio = item.rawData.operatorRatio; // 이미 0~1 사이 값
    const normalizedTime = item.rawData.avgTime > 0 ? 
                           1 - Math.min(item.rawData.avgTime, maxAvgTime) / maxAvgTime : 0; // 소요시간은 짧을수록 좋음
    
    // 가중치 적용 (전역 변수에서 가져옴)
    const scoreComponent = normalizedScore * PERFORMANCE_WEIGHTS.totalScore;
    const ratioComponent = normalizedRatio * PERFORMANCE_WEIGHTS.operatorRatio;
    const timeComponent = normalizedTime * PERFORMANCE_WEIGHTS.avgTime;
    
    // 최종 성과 점수 계산
    const performanceScore = scoreComponent + ratioComponent + timeComponent;
    
    // 디버깅을 위한 로그 출력
    console.log(`담당자: ${item.dispatcher}`);
    console.log(`  총 건수: ${item.total}, 총 점수: ${item.totalScore}, 정규화: ${normalizedScore.toFixed(2)}`);
    console.log(`  오퍼레이터 배차: ${item.operatorDispatch}, 비중: ${(item.operatorRatio * 100).toFixed(2)}%, 정규화: ${normalizedRatio.toFixed(2)}`);
    console.log(`  평균 소요시간: ${item.avgTime.toFixed(2)}초, 정규화: ${normalizedTime.toFixed(2)}`);
    console.log(`  성과 점수 계산: ${scoreComponent.toFixed(2)}(총점${PERFORMANCE_WEIGHTS.totalScore * 100}%) + ${ratioComponent.toFixed(2)}(오퍼${PERFORMANCE_WEIGHTS.operatorRatio * 100}%) + ${timeComponent.toFixed(2)}(시간${PERFORMANCE_WEIGHTS.avgTime * 100}%) = ${performanceScore.toFixed(2)}`);
    
    // 성과 점수 저장
    item.performanceScore = performanceScore;
  });
  
  // 성과 점수 기준 정렬
  dispatchersArray.sort((a, b) => b.performanceScore - a.performanceScore);
  
  // 상위 5명 & 하위 5명 추출
  const topPerformers = dispatchersArray.slice(0, 5);
  
  // 하위 5명 (최소 5명 이상일 경우만)
  const bottomPerformers = dispatchersArray.length > 5 
    ? dispatchersArray.slice(-5).reverse() 
    : [];
  
  return {
    top: topPerformers,
    bottom: bottomPerformers,
    totalCount: dispatchersArray.length
  };
}

// 시간대별 배차 분석을 위한 함수
function getHourlyStats() {
  const { data } = getDispatchData();
  
  // 배분시간 열 인덱스 (17번째 열)
  const distTimeIdx = 17;
  
  // 시간대별 집계
  const hourlyStats = {};
  
  data.forEach(row => {
    const distTimeStr = row[distTimeIdx];
    
    // 날짜 객체로 변환 (올바른 날짜 형식인 경우)
    if (distTimeStr && distTimeStr instanceof Date) {
      const hour = distTimeStr.getHours();
      
      // 시간대가 통계에 없으면 초기화
      if (!hourlyStats[hour]) {
        hourlyStats[hour] = 0;
      }
      
      // 해당 시간대 건수 증가
      hourlyStats[hour]++;
    }
  });
  
  // 시간대 순으로 정렬된 배열로 변환
  const sortedHourlyStats = Object.keys(hourlyStats)
    .map(hour => ({ hour: parseInt(hour), count: hourlyStats[hour] }))
    .sort((a, b) => a.hour - b.hour);
  
  return sortedHourlyStats;
}

// 배차 결과 유형별 통계
function getDispatchResultStats() {
  const { data } = getDispatchData();
  
  // 배차 결과 열 인덱스
  const resultIdx = 5;
  
  // 결과 유형별 집계
  const resultStats = {};
  
  data.forEach(row => {
    const result = row[resultIdx];
    
    // 결과 유형이 통계에 없으면 초기화
    if (!resultStats[result]) {
      resultStats[result] = 0;
    }
    
    // 해당 결과 유형 건수 증가
    resultStats[result]++;
  });
  
  // 건수 기준 내림차순 정렬된 배열로 변환
  const sortedResultStats = Object.keys(resultStats)
    .map(result => ({ result: result, count: resultStats[result] }))
    .sort((a, b) => b.count - a.count);
  
  return sortedResultStats;
}

// 데이터 로드 함수
function loadData(showSpinner = true) {
  // 이전 타이머가 있으면 취소
  if (refreshTimer) {
    clearTimeout(refreshTimer);
  }
  
  // 로딩 중 상태 설정
  isLoading = true;
  
  // 로딩 표시
  if (showSpinner) {
    document.getElementById('loading').style.opacity = '1';
    document.getElementById('dashboard-content').style.opacity = '0.5';
  }
  
  try {
    // 서버에서 데이터 가져오기
    const dispatchData = getDispatchData();
    
    if (!dispatchData || !dispatchData.data || dispatchData.data.length === 0) {
      handleError('데이터를 가져올 수 없습니다.');
      return;
    }
    
    // 전체 요약 통계 계산
    const summary = getSummaryStats();
    
    // 담당자별 통계 계산
    const dispatcherStats = getDispatcherStats();
    
    // 고성과자/저성과자 통계 계산
    const performersData = getTopPerformers();
    
    // 시간대별 통계 계산
    const hourlyStats = getHourlyStats();
    
    // 배차 결과 통계 계산
    const resultStats = getDispatchResultStats();
    
    // 지역별 통계 계산
    const regionData = getRegionStats();
    
    // 데이터 표시
    displaySummary(summary);
    displayDispatcherStats(dispatcherStats);
    displayPerformers(performersData);
    displayHourlyStats(hourlyStats);
    displayDispatchResultStats(resultStats);
    displayRegionChart(regionData);
    
    // 툴팁 업데이트
    updatePerformanceTooltips();
    
    // 로딩 완료 후 차트 컨테이너 크기 조정
    fixChartDisplay();
  } catch (error) {
    console.error('데이터 로딩 중 오류:', error);
    handleError('데이터 로딩 중 오류가 발생했습니다.');
  } finally {
    // 로딩 상태 해제
    isLoading = false;
    
    // 로딩 표시 해제
    if (showSpinner) {
      document.getElementById('loading').style.opacity = '0';
      document.getElementById('dashboard-content').style.opacity = '1';
    }
  }
}