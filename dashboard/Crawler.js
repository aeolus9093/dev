// 담당자별 영역 정의 (시작행, 끝행 형태)
const DISPATCHER_RANGES = [
    [3, 502], [505, 1005], [1008, 1508], [1511, 2011], [2014, 2514], [2517, 3017], [3020, 3520], [3523, 4023], [4026, 4526], [4529, 5029], [5032, 5532], [5535, 6035], [6038, 6538], [6541, 7041], [7044, 7544], [7547, 8047], [8050, 8550], [8553, 9053], [9056, 9556], [9559, 10059], [10062, 10562], [10565, 11065], [11068, 11568], [11571, 12071], [12074, 12574], [12577, 13077], [13080, 13580], [13583, 14083], [14086, 14586], [14589, 15089], [15092, 15592], [15595, 16095], [16098, 16598], [16601, 17101], [17104, 17604], [17607, 18107], [18110, 18610], [18613, 19113], [19116, 19616], [19619, 20119], [20122, 20622], [20625, 21125], [21128, 21628], [21631, 22131], [22134, 22634], [22637, 23137], [23140, 23640], [23643, 24143], [24146, 24646], [24649, 25149], [25152, 25652], [25655, 26155], [26158, 26658], [26661, 27161], [27164, 27664], [27667, 28167], [28170, 28670], [28673, 29173], [29176, 29676], [29679, 30179], [30182, 30682], [30685, 31185], [31188, 31688], [31691, 32191], [32194, 32694]
    ];
    
    /**
     * 메인 데이터 수집 함수
     */
    function crawlDispatchData() {
      // 실행 시간 체크 (새벽 4시~5시 30분 사이에는 실행하지 않음)
      const now = new Date();
      const hour = now.getHours();
      const minute = now.getMinutes();
      
      if (hour === 4 || (hour === 5 && minute < 30)) {
        Logger.log("새벽 4시~5시 30분 사이에는 실행되지 않습니다.");
        return;
      }
      
      // 5시 30분~6시 사이에는 초기화 및 새 탭 생성
      if (hour === 5 && minute >= 30 && minute < 60) {
        resetAndCreateNewTab();
        return;
      }
      
      const sourceSpreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/19m-trjtglnJPiyGe88T66jS-jrGYAF22YYnee1RtYX8/edit?pli=1&gid=2103062223#gid=2103062223");
      const dispatchSheet = sourceSpreadsheet.getSheetByName('배차실적'); // 원본 탭 이름 확인
      const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const resultSheet = currentSpreadsheet.getSheetByName('실적'); // 대상 탭 이름 확인
    
      // 시트가 존재하는지 확인
      if (!dispatchSheet) {
        Logger.log("오류: '배차실적' 시트를 찾을 수 없습니다. 시트 이름을 확인해주세요.");
        return;
      }
    
      if (!resultSheet) {
        Logger.log("오류: '실적' 시트를 찾을 수 없습니다. 필요하다면 createNewResultTab() 함수를 호출하세요.");
        return;
      }
    
      // 스크립트 속성에서 마지막으로 처리한 담당자별 행 번호 가져오기
      const scriptProperties = PropertiesService.getScriptProperties();
      let lastProcessedRows = JSON.parse(scriptProperties.getProperty('lastProcessedRows') || '{}');
    
      // 새로운 데이터를 저장할 배열
      let newData = [];
    
      // 각 담당자별로 데이터 수집
      for (let i = 0; i < DISPATCHER_RANGES.length; i++) {
        const dispatcherId = i + 1; // 담당자 ID (1부터 시작)
        const startRow = DISPATCHER_RANGES[i][0];
        const endRow = DISPATCHER_RANGES[i][1];
    
        // 이전에 처리한 마지막 행 (없으면 시작행 바로 이전으로 설정)
        const lastProcessedRow = lastProcessedRows[dispatcherId] || (startRow - 1);
    
        // 마지막 처리 행 이후부터 종료행까지 검사
        if (lastProcessedRow < endRow) {
          // 검사할 행 범위
          const rangeToCheck = dispatchSheet.getRange(
            lastProcessedRow + 1,
            1,
            endRow - lastProcessedRow,
            19 // A부터 S까지 (19개 열) 가져오기 (실제로는 18열까지만 사용)
          );
    
          const values = rangeToCheck.getValues();
    
          // 현재 담당자 영역에서 발견된 마지막 데이터 행
          let currentLastDataRow = lastProcessedRow;
    
          // 데이터가 있는 행만 필터링하여 newData에 추가
          for (let j = 0; j < values.length; j++) {
            const actualRow = lastProcessedRow + 1 + j; // 실제 시트상의 행 번호
            const row = values[j];
    
            // 행에 데이터가 있는지 확인 (A열 또는 주요 데이터 열에 값이 있는지)
            if (row[0] !== '' || row[1] !== '' || row[2] !== '') {
              // T열은 제외하고 A~S열(0~18)까지만 사용
              const slicedRow = row.slice(0, 19);
              
              // 담당자 ID 추가
              const rowWithDispatcherId = [...slicedRow];
              rowWithDispatcherId.push(dispatcherId);
    
              newData.push(rowWithDispatcherId);
              currentLastDataRow = actualRow; // 발견된 마지막 데이터 행 업데이트
            }
          }
    
          // 마지막으로 처리한 행 업데이트
          lastProcessedRows[dispatcherId] = currentLastDataRow;
        }
      }
    
      // 새로운 데이터가 있으면 실적탭에 추가
      if (newData.length > 0) {
        // 실적탭의 마지막 행을 찾되, T 열에 배열 함수가 존재하므로 A~S 열만 검사
        const dataRange = resultSheet.getRange("A1:S" + resultSheet.getMaxRows());
        const values = dataRange.getValues();
        
        // 데이터가 있는 마지막 행 찾기 (A~S 열 중 하나라도 값이 있는지 확인)
        let lastRow = 0;
        for (let i = values.length - 1; i >= 0; i--) {
          const row = values[i];
          // 행에 데이터가 있는지 확인 (A~S 열)
          const hasData = row.some(cell => cell !== '');
          if (hasData) {
            lastRow = i + 1; // 0-based 인덱스를 1-based 행 번호로 변환
            break;
          }
        }
        
        // 전체 열 수 확인 (담당자 ID를 포함한 전체 열 수)
        const columnCount = newData[0].length;
        
        // 타겟 범위 설정
        const targetRange = resultSheet.getRange(
          lastRow + 1,
          1,
          newData.length,
          columnCount
        );
    
        targetRange.setValues(newData);
    
        Logger.log(`${newData.length}개의 새로운 데이터가 실적탭에 추가되었습니다.`);
      } else {
        Logger.log('새로운 데이터가 없습니다.');
      }
    
      // 마지막으로 처리한 행 저장
      scriptProperties.setProperty('lastProcessedRows', JSON.stringify(lastProcessedRows));
    }
    
    /**
     * 새벽 5:30~6:00 사이에 실행되는 초기화 및 새 탭 생성 함수
     */
    function resetAndCreateNewTab() {
      Logger.log("새벽 5:30~6:00 - 초기화 및 새 탭 생성 작업을 시작합니다.");
      
      const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const resultSheet = currentSpreadsheet.getSheetByName('실적');
      
      if (!resultSheet) {
        Logger.log("'실적' 시트가 없어 초기화 작업을 건너뜁니다.");
        // 새 실적 탭 생성 로직 추가 필요
        return;
      }
      
      try {
        // 1. 날짜 형식으로 이름 변경 (YYYY.MM.DD)
        const yesterday = new Date();
        yesterday.setDate(yesterday.getDate() - 1);
        const formattedDate = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "yyyy.MM.dd");
        resultSheet.setName(formattedDate);
        Logger.log(`기존 '실적' 시트의 이름이 '${formattedDate}'로 변경되었습니다.`);
        
        // 2. 이름이 변경된 시트 복사하여 '실적'이라는 이름으로 생성
        const newSheet = resultSheet.copyTo(currentSpreadsheet).setName('실적');
        Logger.log("새로운 '실적' 시트가 생성되었습니다.");
        
        // A~T 컬럼의 Bold 처리 해제
        newSheet.getRange("A1:T1").setFontWeight("normal");
        Logger.log("A~T 컬럼의 Bold 처리가 해제되었습니다.");
        
        // Q~S 컬럼의 서식을 HH:MM:SS 형태로 변경
        newSheet.getRange("Q:S").setNumberFormat("HH:MM:SS");
        Logger.log("Q~S 컬럼의 서식이 HH:MM:SS 형태로 변경되었습니다.");
        
        // T 컬럼 숨기기 (T는 알파벳 순서로 20번째 열)
        try {
          newSheet.hideColumns(20, 1); // 20번째 열부터 1개 열 숨기기
          Logger.log("T 컬럼(20번째 열)이 숨겨졌습니다.");
        } catch (e) {
          Logger.log("T 컬럼을 숨기는 중 오류가 발생했습니다: " + e.toString());
        }
        
        // 3. 2행부터 마지막 데이터 행까지 전부 제거
        const lastRow = newSheet.getLastRow();
        if (lastRow > 1) {
          newSheet.deleteRows(2, lastRow - 1);
          Logger.log(`2행부터 ${lastRow}행까지 ${lastRow - 1}개 행이 제거되었습니다.`);
        }
        
        // 4. U2 셀에 소요시간 계산 함수 입력
        const arrayFormula = '=ArrayFormula(IF(((S2:S<>"")*(R2:R<>""))*((S2:S-R2:R)*86400>=10)*((S2:S-R2:R)*86400<1200), (S2:S-R2:R)*86400, ""))';
        newSheet.getRange("U2").setFormula(arrayFormula);
        Logger.log("U2 셀에 소요시간 계산 함수가 입력되었습니다.");
        
        // 5. 행 수를 5000개로 설정 (필요시 더 많은 데이터를 담을 수 있도록)
        if (newSheet.getMaxRows() < 5000) {
          const rowsToAdd = 5000 - newSheet.getMaxRows();
          if (rowsToAdd > 0) {
            newSheet.insertRowsAfter(newSheet.getMaxRows(), rowsToAdd);
            Logger.log(`시트의 행 수를 5000개로 늘렸습니다.`);
          }
        }
        
        // 6. 마지막 처리 행 초기화
        initializeLastProcessedRows();
        
        Logger.log("초기화 및 새 탭 생성 작업이 완료되었습니다.");
      } catch (e) {
        Logger.log(`오류 발생: ${e.toString()}`);
      }
    }
    
    /**
     * 시간 기반 트리거 생성 (1분마다)
     */
    function createTimeTrigger() {
      // 기존 트리거 삭제
      const triggers = ScriptApp.getProjectTriggers();
      for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'crawlDispatchData') {
          ScriptApp.deleteTrigger(triggers[i]);
        }
      }
    
      // 새 트리거 생성
      ScriptApp.newTrigger('crawlDispatchData')
        .timeBased()
        .everyMinutes(1)
        .create();
    
      Logger.log('1분마다 실행되는 트리거가 생성되었습니다.');
    }
    
    /**
     * 담당자별 마지막 처리 행 초기화
     */
    function initializeLastProcessedRows() {
      const scriptProperties = PropertiesService.getScriptProperties();
      const lastProcessedRows = {};
    
      for (let i = 0; i < DISPATCHER_RANGES.length; i++) {
        const dispatcherId = i + 1;
        lastProcessedRows[dispatcherId] = DISPATCHER_RANGES[i][0] - 1; // 시작행 바로 이전
      }
    
      scriptProperties.setProperty('lastProcessedRows', JSON.stringify(lastProcessedRows));
      Logger.log('마지막 처리 행이 초기화되었습니다.');
    }
    
    /**
     * 모든 설정을 초기화하고 트리거 생성
     */
    function setupEverything() {
      // 새 실적 탭이 없으면 생성
      const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let resultSheet = currentSpreadsheet.getSheetByName('실적');
      
      if (!resultSheet) {
        createNewResultTab();
      }
      
      initializeLastProcessedRows();
      createTimeTrigger();
      Logger.log('모든 설정이 완료되었습니다.');
    }
