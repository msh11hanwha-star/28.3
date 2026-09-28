/**
 * ☕ 바이브 카페 주문서 스크립트 (script.js)
 * - 실시간 예상 금액 및 주문내역 미리보기 동적 계산
 * - [주문하기] 클릭 시 영수증 표시 및 엑셀 파일(.xlsx) 즉시 자동 다운로드
 * - 이메일 필드 제외 처리
 * - SheetJS 기반 엑셀 파일 생성 (오프라인 시 UTF-8 BOM CSV 자동 백업)
 */

document.addEventListener('DOMContentLoaded', () => {
    // 주요 DOM 요소 선택
    const orderForm = document.getElementById('orderForm');
    const beverageSelect = document.getElementById('userBeverage');
    const menuCards = document.querySelectorAll('.menu-card');
    const sizeRadios = document.querySelectorAll('input[name="size"]');
    const optionCheckboxes = document.querySelectorAll('input[name="options"]');
    const receiptSection = document.getElementById('receiptSection');
    const btnExportExcel = document.getElementById('btnExportExcel');
    const btnScrollTop = document.getElementById('btnScrollTop');

    // 실시간 미리보기 DOM 요소
    const previewBeverage = document.getElementById('previewBeverage');
    const previewSize = document.getElementById('previewSize');
    const previewOptions = document.getElementById('previewOptions');
    const previewTotalPrice = document.getElementById('previewTotalPrice');

    // 최근 주문 정보를 저장할 상태 객체
    let currentOrder = null;

    /* ==========================================================================
       1. 실시간 주문내역 및 예상 결제 금액 계산 함수
       ========================================================================== */
    function updateLivePreview() {
        let beveragePrice = 0;
        let beverageText = '음료를 선택해주세요';

        // 1) 선택된 음료 정보
        if (beverageSelect.value) {
            const selectedOption = beverageSelect.options[beverageSelect.selectedIndex];
            beveragePrice = parseInt(selectedOption.getAttribute('data-price') || '0', 10);
            beverageText = selectedOption.text;
        }

        // 2) 선택된 사이즈 정보
        const selectedSizeInput = document.querySelector('input[name="size"]:checked');
        const sizeName = selectedSizeInput ? selectedSizeInput.getAttribute('data-name') : 'M(보통)';
        const sizePrice = selectedSizeInput ? parseInt(selectedSizeInput.getAttribute('data-price') || '0', 10) : 500;
        const sizeText = `${sizeName} (${sizePrice > 0 ? '+' + sizePrice.toLocaleString() + '원' : '추가금 없음'})`;

        // 3) 선택된 추가 옵션들 정보
        const checkedOptionInputs = document.querySelectorAll('input[name="options"]:checked');
        let optionsPrice = 0;
        const optionsList = [];

        checkedOptionInputs.forEach(input => {
            const price = parseInt(input.getAttribute('data-price') || '0', 10);
            optionsPrice += price;
            optionsList.push(`${input.value}(+${price.toLocaleString()}원)`);
        });

        const optionsText = optionsList.length > 0 ? optionsList.join(', ') : '선택 없음';

        // 4) 총 예상 금액 계산 (음료가 선택되지 않았을 때는 0원으로 표시)
        const totalPrice = beverageSelect.value ? (beveragePrice + sizePrice + optionsPrice) : 0;

        // 5) 화면 렌더링
        if (previewBeverage) previewBeverage.textContent = beverageText;
        if (previewSize) previewSize.textContent = sizeText;
        if (previewOptions) previewOptions.textContent = optionsText;
        if (previewTotalPrice) previewTotalPrice.textContent = `${totalPrice.toLocaleString()}원`;

        return {
            beveragePrice,
            beverageText,
            sizeName,
            sizePrice,
            sizeText,
            optionsList,
            optionsText,
            optionsPrice,
            totalPrice
        };
    }

    /* ==========================================================================
       2. 메뉴 이미지 카드와 드롭다운 양방향 연동
       ========================================================================== */
    // 카드 클릭 시 드롭다운 값 변경, active 클래스 갱신 및 실시간 금액 업데이트
    menuCards.forEach(card => {
        card.addEventListener('click', () => {
            const beverageValue = card.getAttribute('data-value');
            beverageSelect.value = beverageValue;

            menuCards.forEach(c => c.classList.remove('active'));
            card.classList.add('active');

            updateLivePreview();
        });
    });

    // 드롭다운 변경 시 메뉴 카드 동기화 및 실시간 금액 업데이트
    beverageSelect.addEventListener('change', () => {
        const selectedValue = beverageSelect.value;
        menuCards.forEach(card => {
            if (card.getAttribute('data-value') === selectedValue) {
                card.classList.add('active');
            } else {
                card.classList.remove('active');
            }
        });
        updateLivePreview();
    });

    // 사이즈 및 추가 옵션 변경 시 실시간 금액 즉시 업데이트
    sizeRadios.forEach(radio => radio.addEventListener('change', updateLivePreview));
    optionCheckboxes.forEach(checkbox => checkbox.addEventListener('change', updateLivePreview));

    /* ==========================================================================
       3. 폼 초기화 (다시 작성 버튼)
       ========================================================================== */
    orderForm.addEventListener('reset', () => {
        menuCards.forEach(card => card.classList.remove('active'));
        receiptSection.style.display = 'none';
        currentOrder = null;

        // 초기화 후 기본 상태(M 사이즈 기본값 등)로 실시간 안내 갱신
        setTimeout(updateLivePreview, 10);
    });

    /* ==========================================================================
       4. 주문하기 버튼 클릭: 영수증 표시 & 엑셀 파일 즉시 자동 다운로드
       ========================================================================== */
    orderForm.addEventListener('submit', (e) => {
        e.preventDefault();

        // 음료 선택 여부 검증
        if (!beverageSelect.value) {
            alert('음료를 선택해주세요.');
            beverageSelect.focus();
            return;
        }

        // 1) 기본 정보 수집 (이메일 제외)
        const userName = document.getElementById('userName').value.trim();
        const userPhone = document.getElementById('userPhone').value.trim() || '미입력';
        const userRequest = document.getElementById('userRequest').value.trim() || '없음';

        // 2) 계산된 주문 내역 가져오기
        const calc = updateLivePreview();

        // 3) 주문 일시 문자열 생성
        const now = new Date();
        const dateString = now.toLocaleString('ko-KR', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit'
        });

        // 4) 현재 주문 객체 구성 (이메일 없음)
        currentOrder = {
            orderDate: dateString,
            name: userName,
            phone: userPhone,
            beverage: calc.beverageText,
            size: calc.sizeText,
            options: calc.optionsText,
            request: userRequest,
            totalPrice: calc.totalPrice,
            formattedTotal: `${calc.totalPrice.toLocaleString()}원`
        };

        // 5) 영수증 화면에 데이터 렌더링
        document.getElementById('receiptDate').textContent = `주문 일시: ${dateString}`;
        document.getElementById('receiptName').textContent = userName;
        document.getElementById('receiptPhone').textContent = userPhone;
        document.getElementById('receiptBeverage').textContent = currentOrder.beverage;
        document.getElementById('receiptSize').textContent = currentOrder.size;
        document.getElementById('receiptOptions').textContent = currentOrder.options;
        document.getElementById('receiptRequest').textContent = currentOrder.request;
        document.getElementById('receiptTotalPrice').textContent = currentOrder.formattedTotal;

        // 6) 영수증 영역 보이기 & 부드럽게 스크롤
        receiptSection.style.display = 'block';
        receiptSection.scrollIntoView({ behavior: 'smooth' });

        // 7) [요청 반영] 주문하기 버튼을 누르면 그 즉시 엑셀 파일 자동 다운로드 실행!
        exportOrderToExcel(currentOrder);
    });

    /* ==========================================================================
       5. 엑셀 파일(.xlsx) 생성 및 다운로드 함수
       ========================================================================== */
    function exportOrderToExcel(order) {
        if (!order) return;

        const today = new Date().toISOString().slice(0, 10);
        const fileName = `바이브카페_주문서_${order.name}_${today}`;

        // SheetJS (XLSX) 사용 가능한 경우 (.xlsx 파일 생성)
        if (typeof XLSX !== 'undefined') {
            const excelData = [
                ['바이브 카페 주문서'],
                ['주문 일시', order.orderDate],
                [],
                ['항목', '주문 내용'],
                ['주문자 이름', order.name],
                ['전화번호', order.phone],
                ['선택 음료', order.beverage],
                ['사이즈', order.size],
                ['추가 옵션', order.options],
                ['요청사항', order.request],
                ['총 결제금액', order.formattedTotal]
            ];

            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(excelData);

            // 열 너비 지정
            ws['!cols'] = [
                { wch: 15 },
                { wch: 35 }
            ];

            XLSX.utils.book_append_sheet(wb, ws, '주문서');
            XLSX.writeFile(wb, `${fileName}.xlsx`);
        } else {
            // SheetJS 로드 불가 시 UTF-8 BOM CSV 파일로 자동 백업
            exportBackupCSV(order, fileName);
        }
    }

    /**
     * 오프라인 시 한글 깨짐 없는 CSV(UTF-8 BOM) 파일 다운로드 백업 함수
     */
    function exportBackupCSV(order, fileName) {
        const rows = [
            ['항목', '주문 내용'],
            ['주문 일시', `"${order.orderDate}"`],
            ['주문자 이름', `"${order.name}"`],
            ['전화번호', `"${order.phone}"`],
            ['선택 음료', `"${order.beverage}"`],
            ['사이즈', `"${order.size}"`],
            ['추가 옵션', `"${order.options}"`],
            ['요청사항', `"${order.request.replace(/"/g, '""')}"`],
            ['총 결제금액', `"${order.formattedTotal}"`]
        ];

        const csvContent = '\uFEFF' + rows.map(r => r.join(',')).join('\r\n');
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const downloadUrl = URL.createObjectURL(blob);
        const link = document.createElement('a');

        link.href = downloadUrl;
        link.download = `${fileName}.csv`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(downloadUrl);
    }

    // 영수증 하단 '엑셀 파일 다시 다운로드' 버튼 클릭 이벤트
    if (btnExportExcel) {
        btnExportExcel.addEventListener('click', () => {
            if (currentOrder) {
                exportOrderToExcel(currentOrder);
            } else {
                alert('주문 내역이 없습니다.');
            }
        });
    }

    /* ==========================================================================
       6. 플로팅 스크롤 탑 버튼 (Circular floating control)
       ========================================================================== */
    if (btnScrollTop) {
        window.addEventListener('scroll', () => {
            if (window.scrollY > 200) {
                btnScrollTop.style.display = 'flex';
            } else {
                btnScrollTop.style.display = 'none';
            }
        });

        btnScrollTop.addEventListener('click', () => {
            window.scrollTo({
                top: 0,
                behavior: 'smooth'
            });
        });
    }

    // 최초 로드 시 실시간 안내 초기 렌더링
    updateLivePreview();
});
