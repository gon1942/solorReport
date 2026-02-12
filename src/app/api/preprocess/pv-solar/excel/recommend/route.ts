import { NextRequest, NextResponse } from 'next/server';
import ExcelJS from 'exceljs';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;
    const sheetsData = JSON.parse(formData.get('sheets') as string);

    if (!file || !sheetsData) {
      return NextResponse.json(
        { success: false, error: 'File and sheets data required' },
        { status: 400 }
      );
    }

    // Excel 파일 분석
    const buffer = Buffer.from(await file.arrayBuffer()) as any;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    // PV Solar 특화 AI 추천 로직
    const recommendations = await generatePVSolarRecommendations(sheetsData, workbook);

    console.log('✅ PV Solar AI 추천 결과:', {
      선택시트: recommendations.selectedSheet,
      추천열수: recommendations.columns?.length || 0,
      추천열목록: recommendations.columns,
      신뢰도: Math.round((recommendations.confidence || 0) * 100) + '%',
      AI제목: recommendations.aiGeneratedTitle,
      AI설명길이: recommendations.aiGeneratedDescription?.length || 0,
      데이터카테고리: recommendations.dataCategory,
      핵심인사이트수: recommendations.keyInsights?.length || 0,
      pvSolar인사이트: recommendations.pvSolarInsights
    });

    return NextResponse.json({
      success: true,
      data: recommendations
    });

  } catch (error) {
    console.error('PV Solar Excel AI recommendation error:', error);
    return NextResponse.json(
      { 
        success: false, 
        error: error instanceof Error ? error.message : 'PV Solar AI recommendation failed' 
      },
      { status: 500 }
    );
  }
}

async function generatePVSolarRecommendations(sheetsData: any[], workbook: ExcelJS.Workbook) {
  const recommendations: any = {
    selectedSheet: '',
    columns: [],
    structure: '',
    description: '',
    confidence: 0,
    aiGeneratedTitle: '',
    aiGeneratedDescription: '',
    dataCategory: 'PV Solar',
    keyInsights: [],
    pvSolarInsights: {
      peakProductionTimes: [],
      weatherFactors: [],
      maintenancePatterns: [],
      performanceMetrics: []
    }
  };

  // 1. PV Solar 적합한 시트 선택
  let bestSheet = null;
  let bestScore = 0;

  for (const sheet of sheetsData) {
    let score = 0;
    
    // 데이터 양 점수 (더 많은 데이터가 있는 시트 선호)
    score += Math.min(sheet.sampleData?.length || 0, 10) * 10;
    
    // 헤더 품질 점수 (빈 헤더가 적은 시트 선호)
    const nonEmptyHeaders = sheet.headers?.filter((h: string) => h && h.trim()) || [];
    score += nonEmptyHeaders.length * 5;
    
    // PV Solar 관련 키워드 점수
    const headerText = nonEmptyHeaders.join(' ').toLowerCase();
    if (headerText.includes('energy') || headerText.includes('power') || headerText.includes('production')) score += 30;
    if (headerText.includes('solar') || headerText.includes('pv') || headerText.includes('panel')) score += 25;
    if (headerText.includes('weather') || headerText.includes('temperature') || headerText.includes('irradiance')) score += 20;
    if (headerText.includes('performance') || headerText.includes('efficiency') || headerText.includes('output')) score += 15;
    
    // 첫 번째 시트 보너스
    if (sheetsData.indexOf(sheet) === 0) {
      score += 10;
    }

    if (score > bestScore) {
      bestScore = score;
      bestSheet = sheet;
    }
  }

  if (bestSheet) {
    recommendations.selectedSheet = bestSheet.name;
    recommendations.confidence = Math.min(bestScore / 100, 1);

    // 2. PV Solar 기반 AI 분석
    const headers = bestSheet.headers || [];
    const sampleData = bestSheet.sampleData || [];
    const actualRowCount = bestSheet.rowCount ? bestSheet.rowCount - 1 : (sampleData.length - 1);
    const actualColumnCount = bestSheet.columnCount || headers.length;
    
    const pvSolarAnalysis = await performPVSolarAnalysis(bestSheet.name, headers, sampleData, actualRowCount, actualColumnCount);
    
    recommendations.aiGeneratedTitle = pvSolarAnalysis.title;
    recommendations.aiGeneratedDescription = pvSolarAnalysis.description;
    recommendations.dataCategory = 'PV Solar';
    recommendations.keyInsights = pvSolarAnalysis.insights;
    recommendations.pvSolarInsights = pvSolarAnalysis.pvSolarInsights;
    
    // 3. PV Solar 특화 필드 추천
    if (pvSolarAnalysis.recommendedFields && pvSolarAnalysis.recommendedFields.length > 0) {
      recommendations.columns = pvSolarAnalysis.recommendedFields.filter((fieldName: any) =>
        headers.includes(fieldName)
      );
      
      if (recommendations.columns.length === 0) {
        console.warn('⚠️ PV Solar 추천 필드가 실제 헤더와 매치되지 않음, 상위 필드로 대체');
        recommendations.columns = headers.slice(0, Math.min(5, headers.length));
      }
    } else {
      console.warn('⚠️ PV Solar AI에서 추천 필드를 반환하지 않음, 상위 필드로 대체');
      recommendations.columns = headers.slice(0, Math.min(5, headers.length));
    }

    // 4. PV Solar 구조 설명 생성
    const totalRows = actualRowCount;
    const totalCols = actualColumnCount;
    
    recommendations.structure = `PV Solar 데이터 (${totalRows}행 × ${totalCols}열)`;
    recommendations.description = 
      `PV Solar 데이터 전문 AI가 시트 '${bestSheet.name}'에서 ${recommendations.columns.length}개의 핵심 필드를 선별했습니다. ` +
      `${totalRows}행 × ${totalCols}열 규모의 에너지 생산 데이터를 분석하여 ` +
      `최적의 데이터 구조를 구축합니다.`;

    // 5. PV Solar 특화 처리 옵션 추천
    recommendations.processing = {
      headerNormalization: true,
      preserveFormatting: true,
      suggestedFilters: []
    };

    // PV Solar 특화 필터 추천
    const selectedHeaders = recommendations.columns;
    const dateColumns = selectedHeaders.filter((header: any) =>
      header.toLowerCase().includes('date') || header.toLowerCase().includes('날짜') || 
      header.toLowerCase().includes('time') || header.toLowerCase().includes('시간')
    );
    const energyColumns = selectedHeaders.filter((header: any) =>
      header.toLowerCase().includes('energy') || header.toLowerCase().includes('power') || 
      header.toLowerCase().includes('production') || header.toLowerCase().includes('output') ||
      header.toLowerCase().includes('에너지') || header.toLowerCase().includes('전력') ||
      header.toLowerCase().includes('발전량') || header.toLowerCase().includes('생산량')
    );
    const weatherColumns = selectedHeaders.filter((header: any) =>
      header.toLowerCase().includes('weather') || header.toLowerCase().includes('temperature') || 
      header.toLowerCase().includes('irradiance') || header.toLowerCase().includes('solar') ||
      header.toLowerCase().includes('날씨') || header.toLowerCase().includes('기온') ||
      header.toLowerCase().includes('일사량') || header.toLowerCase().includes('태양')
    );

    if (dateColumns.length > 0) {
      recommendations.processing.suggestedFilters.push({
        type: 'date_range',
        column: dateColumns[0],
        description: `${dateColumns[0]} 기준으로 최신 데이터 우선 필터링 (PV 데이터 최신성 중요)`
      });
    }

    if (energyColumns.length > 0) {
      recommendations.processing.suggestedFilters.push({
        type: 'numeric_range',
        column: energyColumns[0],
        description: `${energyColumns[0]} 유효값 범위 필터링 (에너지 데이터 품질 관리)`
      });
    }

    if (weatherColumns.length > 0) {
      recommendations.processing.suggestedFilters.push({
        type: 'weather_range',
        column: weatherColumns[0],
        description: `${weatherColumns[0]} 데이터 필터링 (날씨 데이터의 영향력 분석)`
      });
    }
  }

  return recommendations;
}

async function performPVSolarAnalysis(sheetName: string, headers: string[], sampleData: any[][], actualRowCount: number, actualColumnCount: number) {
  const pvSolarDataAnalysis = analyzePVSolarDataPatterns(headers, sampleData, actualRowCount);

  try {
    // PV Solar 특화 AI 분석 프롬프트
    const pvSolarPrompt = createPVSolarAnalysisPrompt(sheetName, headers, sampleData, pvSolarDataAnalysis, actualRowCount, actualColumnCount);
    
    const aiResponse = await callAIProvider(pvSolarPrompt);
    
    if (aiResponse.success) {
      return aiResponse.data;
    } else {
      // AI 호출 실패 시 PV Solar 규칙 기반 분석 사용
      return {
        title: generatePVSolarTitle(sheetName, headers, pvSolarDataAnalysis),
        description: generatePVSolarDescription(headers, sampleData, pvSolarDataAnalysis, actualRowCount, actualColumnCount),
        category: 'PV Solar',
        insights: generatePVSolarInsights(headers, sampleData, pvSolarDataAnalysis, actualRowCount),
        recommendedFields: getRecommendedPVFields(headers, pvSolarDataAnalysis),
        pvSolarInsights: generatePVSolarSpecificInsights(pvSolarDataAnalysis)
      };
    }
  } catch (error) {
    console.error('PV Solar AI 분석 실패:', error);
    return {
      title: generatePVSolarTitle(sheetName, headers, pvSolarDataAnalysis),
      description: generatePVSolarDescription(headers, sampleData, pvSolarDataAnalysis, actualRowCount, actualColumnCount),
      category: 'PV Solar',
      insights: generatePVSolarInsights(headers, sampleData, pvSolarDataAnalysis, actualRowCount),
      recommendedFields: getRecommendedPVFields(headers, pvSolarDataAnalysis),
      pvSolarInsights: generatePVSolarSpecificInsights(pvSolarDataAnalysis)
    };
  }
}

function analyzePVSolarDataPatterns(headers: string[], sampleData: any[][], actualRowCount?: number) {
  const patterns = {
    hasEnergyData: false,
    hasProductionData: false,
    hasWeatherData: false,
    hasTimeData: false,
    hasLocationData: false,
    hasPerformanceData: false,
    dataTypes: {} as Record<string, string>,
    recordCount: actualRowCount !== undefined ? actualRowCount : (sampleData.length - 1),
    energyUnits: new Set<string>(),
    timeFormats: new Set<string>()
  };

  headers.forEach((header, index) => {
    const headerLower = header.toLowerCase();
    const samples = sampleData.slice(1).map(row => row[index]).filter(val => val);
    
    // 에너지 데이터 감지
    if (headerLower.includes('energy') || headerLower.includes('power') || 
        headerLower.includes('production') || headerLower.includes('output') ||
        headerLower.includes('발전량') || headerLower.includes('전력') ||
        headerLower.includes('에너지') || headerLower.includes('생산량')) {
      patterns.hasEnergyData = true;
      // 에너지 단위 감지
      samples.forEach(sample => {
        const sampleStr = String(sample).toLowerCase();
        if (sampleStr.includes('kw')) patterns.energyUnits.add('kW');
        if (sampleStr.includes('mw')) patterns.energyUnits.add('MW');
        if (sampleStr.includes('wh')) patterns.energyUnits.add('Wh');
        if (sampleStr.includes('mwh')) patterns.energyUnits.add('MWh');
      });
    }
    
    // 생산 데이터 감지
    if (headerLower.includes('production') || headerLower.includes('generation') ||
        headerLower.includes('발전') || headerLower.includes('생산')) {
      patterns.hasProductionData = true;
    }
    
    // 날씨 데이터 감지
    if (headerLower.includes('weather') || headerLower.includes('temperature') ||
        headerLower.includes('irradiance') || headerLower.includes('solar') ||
        headerLower.includes('일사량') || headerLower.includes('기온')) {
      patterns.hasWeatherData = true;
    }
    
    // 시간 데이터 감지
    if (headerLower.includes('time') || headerLower.includes('date') ||
        headerLower.includes('hour') || headerLower.includes('timestamp') ||
        headerLower.includes('시간') || headerLower.includes('날짜')) {
      patterns.hasTimeData = true;
      // 시간 형식 감지
      samples.forEach(sample => {
        const sampleStr = String(sample);
        if (sampleStr.includes(':')) patterns.timeFormats.add('HH:mm');
        if (sampleStr.includes('-')) patterns.timeFormats.add('YYYY-MM-DD');
        if (sampleStr.includes('/')) patterns.timeFormats.add('YYYY/MM/DD');
      });
    }
    
    // 위치 데이터 감지
    if (headerLower.includes('location') || headerLower.includes('address') ||
        headerLower.includes('city') || headerLower.includes('region') ||
        headerLower.includes('위치') || headerLower.includes('주소') ||
        headerLower.includes('지역')) {
      patterns.hasLocationData = true;
    }
    
    // 성능 데이터 감지
    if (headerLower.includes('performance') || headerLower.includes('efficiency') ||
        headerLower.includes('efficiency') || headerLower.includes('performance') ||
        headerLower.includes('효율') || headerLower.includes('성능')) {
      patterns.hasPerformanceData = true;
    }
    
    // 데이터 타입 분석
    if (samples.length > 0) {
      const sample = samples[0];
      if (!isNaN(Number(sample)) && sample !== '') {
        patterns.dataTypes[header] = 'numeric';
      } else if (String(sample).includes('@')) {
        patterns.dataTypes[header] = 'email';
      } else if (String(sample).includes('-') && String(sample).length >= 8) {
        patterns.dataTypes[header] = 'date';
      } else {
        patterns.dataTypes[header] = 'text';
      }
    }
  });

  return patterns;
}

function createPVSolarAnalysisPrompt(sheetName: string, headers: string[], sampleData: any[][], patterns: any, actualRowCount?: number, actualColumnCount?: number) {
  const sampleRows = sampleData.slice(1, Math.min(6, sampleData.length));
  
  return `PV Solar 데이터 분석 전문가로서 다음 데이터를 분석하고 JSON 형태로 응답해주세요.

**PV Solar 데이터 정보:**
- 시트명: ${sheetName}
- 총 데이터 행수: ${actualRowCount !== undefined ? actualRowCount : patterns.recordCount}개 (헤더 제외)
- 총 열수: ${actualColumnCount !== undefined ? actualColumnCount : headers.length}개
- 전체 열 목록: ${headers.join(', ')}

**데이터 특성 분석:**
- 에너지 데이터 포함: ${patterns.hasEnergyData ? '예' : '아니오'}
- 생산 데이터 포함: ${patterns.hasProductionData ? '예' : '아니오'}
- 날씨 데이터 포함: ${patterns.hasWeatherData ? '예' : '아니오'}
- 시간 데이터 포함: ${patterns.hasTimeData ? '예' : '아니오'}
- 위치 데이터 포함: ${patterns.hasLocationData ? '예' : '아니오'}
- 성능 데이터 포함: ${patterns.hasPerformanceData ? '예' : '아니오'}
- 감지된 에너지 단위: ${Array.from(patterns.energyUnits).join(', ') || '정보 없음'}

**샘플 데이터:**
${sampleRows.map((row, idx) => 
  `행 ${idx + 1}: ${row.map((cell, cellIdx) => `${headers[cellIdx]}: ${cell}`).join(' | ')}`
).join('\n')}

**PV 데이터 분석 지침:**
1. **필드 선별 기준**: PV Solar 데이터 분석과 성능 최적화에 필수적인 필드를 5-8개만 엄선해주세요.
   - 핵심 필드: 발전량, 전력 출력, 시간 기록, 날씨 정보, 위치 정보
   - 성능 관련: 효율률, 정격 출력, 실측 값
   - 환경 관련: 일사량, 기온, 습도, 풍속

2. **제외 필드**:
   - 중복된 필드 (예: timestamp와 time이 둘 다 있으면 timestamp 우선)
   - 메타데이터: created_at, updated_at, id 등
   - 의미 없는 값이 많은 필드
   - 분석과 무관한 관리 필드

3. **PV Solar 특화 분석 요청**:
   - 생산 피크 시간대 분석
   - 날씨 요인별 생산량 영향 분석
   - 유지보수 패턴 감지
   - 성능 메트릭 추출 (실측/정격 비율 등)

다음 형식의 JSON으로 응답해주세요:
{
  "title": "PV Solar [프로젝트/설비명] 데이터",
  "description": "PV Solar 데이터의 특성, 규모, 분석 가능성에 대한 설명",
  "insights": ["데이터에서 발견할 수 있는 핵심 인사이트 1", "비즈니스에 활용 가능한 인사이트 2", "데이터 품질이나 특성에 대한 인사이트 3"],
  "recommendedFields": ["필드명1", "필드명2", "필드명3", "필드명4", "필드명5"],
  "pvSolarInsights": {
    "peakProductionTimes": ["예상 생산 피크 시간대"],
    "weatherFactors": ["날씨 요인"],
    "maintenancePatterns": ["유지보수 패턴"],
    "performanceMetrics": ["성능 메트릭"]
  }
}

필드명은 실제 헤더 목록에서 정확히 사용된 이름을 그대로 사용해주세요.`;
}

function generatePVSolarTitle(sheetName: string, headers: string[], patterns: any) {
  if (patterns.hasEnergyData && patterns.hasTimeData) {
    return 'PV Solar 발전 데이터';
  } else if (patterns.hasPerformanceData) {
    return 'PV Solar 성능 모니터링 데이터';
  } else if (patterns.hasWeatherData) {
    return 'PV Solar 환경 데이터';
  } else {
    return `PV Solar ${sheetName} 데이터`;
  }
}

function generatePVSolarDescription(headers: string[], sampleData: any[][], patterns: any, actualRowCount?: number, actualColumnCount?: number) {
  const recordCount = actualRowCount !== undefined ? actualRowCount : patterns.recordCount;
  const columnCount = actualColumnCount !== undefined ? actualColumnCount : headers.length;
  
  let description = `총 ${recordCount}건의 PV Solar 데이터로 구성된 ${columnCount}개 필드의 구조화된 데이터입니다.`;
  
  const features = [];
  if (patterns.hasEnergyData) features.push('에너지 생산');
  if (patterns.hasWeatherData) features.push('날씨/환경');
  if (patterns.hasTimeData) features.push('시간 기록');
  if (patterns.hasLocationData) features.push('위치 정보');
  if (patterns.hasPerformanceData) features.push('성능 모니터링');
  
  if (features.length > 0) {
    description += `${features.join(', ')} 데이터를 포함하여 종합적인 PV 시스템 분석이 가능합니다.`;
  }
  
  if (patterns.hasEnergyData && patterns.hasWeatherData) {
    description += ' 날씨 요인별 발전량 분석, 시간대별 생산 패턴 분석이 가능합니다.';
  }
  
  return description;
}

function generatePVSolarInsights(headers: string[], sampleData: any[][], patterns: any, actualRowCount?: number) {
  const insights = [];
  
  const recordCount = actualRowCount !== undefined ? actualRowCount : patterns.recordCount;
  
  if (patterns.hasEnergyData && patterns.hasTimeData && recordCount > 100) {
    insights.push('시계열 분석을 통한 일별/주별 발전 트렌드 파악 가능');
  }
  
  if (patterns.hasEnergyData && patterns.hasWeatherData) {
    insights.push('날씨 요인과 발전량 상관관계 분석으로 최적 운전 조건 도출');
  }
  
  if (patterns.hasPerformanceData && recordCount > 50) {
    insights.push('성능 모니터링을 통한 효율 저하 요인 분석 가능');
  }
  
  const energyFieldsCount = Object.values(patterns.dataTypes).filter(type => type === 'numeric').length;
  if (energyFieldsCount > 2) {
    insights.push('다양한 수치 데이터를 활용한 통계 분석 및 예측 모델링 가능');
  }
  
  return insights;
}

function getRecommendedPVFields(headers: string[], patterns: any) {
  const recommended = [];
  
  // 핵심 필드 우선 선택
  if (patterns.hasEnergyData) {
    const energyFields = headers.filter(h => 
      h.toLowerCase().includes('energy') || h.toLowerCase().includes('power') ||
      h.toLowerCase().includes('production') || h.toLowerCase().includes('output') ||
      h.toLowerCase().includes('발전량') || h.toLowerCase().includes('전력') ||
      h.toLowerCase().includes('에너지') || h.toLowerCase().includes('생산량')
    );
    recommended.push(...energyFields.slice(0, 2));
  }
  
  // 시간 필드
  if (patterns.hasTimeData) {
    const timeFields = headers.filter(h => 
      h.toLowerCase().includes('time') || h.toLowerCase().includes('date') ||
      h.toLowerCase().includes('timestamp') ||
      h.toLowerCase().includes('시간') || h.toLowerCase().includes('날짜')
    );
    recommended.push(...timeFields.slice(0, 1));
  }
  
  // 날씨 필드
  if (patterns.hasWeatherData) {
    const weatherFields = headers.filter(h => 
      h.toLowerCase().includes('weather') || h.toLowerCase().includes('temperature') ||
      h.toLowerCase().includes('irradiance') ||
      h.toLowerCase().includes('일사량') || h.toLowerCase().includes('기온')
    );
    recommended.push(...weatherFields.slice(0, 1));
  }
  
  // 성능 필드
  if (patterns.hasPerformanceData) {
    const performanceFields = headers.filter(h => 
      h.toLowerCase().includes('performance') || h.toLowerCase().includes('efficiency') ||
      h.toLowerCase().includes('성능') || h.toLowerCase().includes('효율')
    );
    recommended.push(...performanceFields.slice(0, 1));
  }
  
  // 중복 제거 후 상위 5개 필드 반환
  return [...new Set(recommended)].slice(0, 5);
}

function generatePVSolarSpecificInsights(patterns: any) {
  const insights = {
    peakProductionTimes: patterns.hasTimeData ? ['10:00-14:00'] : [],
    weatherFactors: patterns.hasWeatherData ? ['일사량', '기온'] : [],
    maintenancePatterns: [],
    performanceMetrics: []
  };
  
  if (patterns.hasEnergyData && patterns.hasTimeData) {
    insights.peakProductionTimes = ['08:00-16:00', '12:00-14:00 (피크 타임)'];
  }
  
  if (patterns.hasWeatherData) {
    insights.weatherFactors = ['일사량', '기온', '운량'];
  }
  
  // if (patterns.hasPerformanceData) {
  //   insights.performanceMetrics = ['실측/정격 비율', '시스템 효율', '에너지 품질'];
  // }
  
  return insights;
}

// AI 프로바이더 호출 함수 (범용 AI 엔드포인트 사용)
async function callAIProvider(prompt: string) {
  try {
    const { getApiServerUrl } = await import('@/config/serverConfig');
    const apiServerUrl = getApiServerUrl();
    
    const response = await fetch(`${apiServerUrl}/callai`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        messages: [{ role: 'user', content: prompt }],
        options: {
          provider: 'ollama',
          model: 'airun-chat:latest',
          temperature: 0.1,
          max_tokens: 4000
        }
      })
    });

    if (!response.ok) {
      throw new Error(`AI API 호출 실패: ${response.status} ${response.statusText}`);
    }

    const result = await response.json();
    
    if (result?.success && result?.data?.response) {
      const aiResponse = result.data.response.content || result.data.response;
      
      try {
        const parsedData = JSON.parse(aiResponse);
        return {
          success: true,
          data: parsedData
        };
      } catch (parseError) {
        const jsonMatch = aiResponse.match(/```json\s*([\s\S]*?)\s*```/) || aiResponse.match(/\{[\s\S]*\}/);
        
        if (jsonMatch) {
          const jsonStr = jsonMatch[1] || jsonMatch[0];
          try {
            const parsedData = JSON.parse(jsonStr);
            return {
              success: true,
              data: parsedData
            };
          } catch (regexParseError) {
            // JSON 파싱 실패
          }
        }
      }
    }
    
    throw new Error('AI 응답에서 유효한 JSON을 찾을 수 없음');
    
  } catch (error) {
    console.error('❌ 범용 AI 호출 오류:', error);
    return {
      success: false,
      error: error instanceof Error ? error.message : 'Unknown error'
    };
  }
}