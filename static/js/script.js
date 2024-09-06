let trailers = [];

        function loadTrailers() {
            $.ajax({
                url: '/trailers',
                type: 'GET',
                success: function(data) {
                    trailers = data;
                    const trailerSelect = $('#trailer');
                    trailerSelect.empty();
                    trailerSelect.append('<option value="">Выбрать:</option>');
                    trailers.forEach(trailer => {
                        trailerSelect.append(`<option value="${trailer.id}">${trailer.number}</option>`);
                    });
                    trailerSelect.select2({ tags: true });
                },
                error: function() {
                    alert('Ошибка при загрузке данных о прицепах');
                }
            });
        }

        function calculateTotalWeight() {
            let totalWeight = 0;
            const inputs = document.querySelectorAll('#sectionsContainer input[type="number"]');
            inputs.forEach(input => {
                const value = parseFloat(input.value) || 0; // Если не число, то 0
                totalWeight += value;
            });
            document.getElementById('physical_weight').value = totalWeight; // Обновляем поле
        }

        function updateSections() {
            const select = document.getElementById('trailer');
            const selectedTrailerId = select.value;
            const selectedTrailer = trailers.find(trailer => trailer.id == selectedTrailerId);
            const container = document.getElementById('sectionsContainer');
            container.innerHTML = '';

            if (selectedTrailer) {
                const sections = selectedTrailer.sections.filter(weight => weight !== null); // Фильтруем секции
                container.style.display = sections.length > 0 ? 'block' : 'none'; // Показываем или скрываем контейнер

                sections.forEach((weight, index) => {
                    const div = document.createElement('div');
                    div.className = 'section';

                    const label = document.createElement('label');
                    label.innerText = `Секция ${index + 1}: `;
                    div.appendChild(label);

                    const input = document.createElement('input');
                    input.type = 'number';
                    input.name = `section_weight_${index + 1}`;
                    input.placeholder = `Максимальный вес: ${weight}`;

                    const buttonSet = document.createElement('button');
                    buttonSet.type = 'button';
                    buttonSet.innerText = 'Заполнить до макс';
                    buttonSet.addEventListener('click', function() {
                        input.value = weight; // Заполняем инпут максимальным весом
                        input.dispatchEvent(new Event('input')); // Обновляем цвет
                    });

                    input.addEventListener('input', function() {
                        const value = parseFloat(this.value);
                        if (!isNaN(value)) {
                            this.style.color = (value > weight) ? 'red' : (value < weight) ? 'blue' : 'green';
                            calculateTotalWeight();
                        } else {
                            this.style.color = 'black';
                        }
                    });

                    div.appendChild(input);
                    div.appendChild(buttonSet); // Кнопка для заполнения

                    // Добавляем поля для уникальных значений
                    const attributes = [
                        { label: 'Массовая доля жира %', name: `fat_content_${index + 1}` },
                        { label: 'Массовая доля белка %', name: `protein_content_${index + 1}` },
                        { label: 'Кислотность °Т', name: `acidity_${index + 1}` },
                        { label: 'Температура °С', name: `temperature_${index + 1}` },
                        { label: 'Плотность кг/м3', name: `density_${index + 1}` },
                        { label: 'Содер. Самат. Клеток, тыс/см3', name: `cell_content_${index + 1}` },
                        { label: 'Группа чистоты', name: `purity_group_${index + 1}` },
                        { label: 'Термоустойчивочть, группа', name: `heat_resistance_${index + 1}` },
                        { label: 'Сорт', name: `grade_${index + 1}` }
                    ];

                    const attributesContainer = document.createElement('div');
                    attributesContainer.className = 'attributes-container'; // Новый контейнер для атрибутов

                    attributes.forEach(attr => {
                        const attrDiv = document.createElement('div');
                        attrDiv.className = 'attribute';

                        const attrInput = document.createElement('input');
                        attrInput.type = 'text';
                        attrInput.name = attr.name;
                        attrInput.placeholder = attr.label; // Подсказка в поле ввода

                        attrDiv.appendChild(attrInput);
                        attributesContainer.appendChild(attrDiv); // Добавляем атрибут в контейнер
                    });

                    div.appendChild(attributesContainer); // Добавляем контейнер атрибутов в секцию
                    container.appendChild(div); // Добавляем секцию в контейнер
                });
            } else {
                container.style.display = 'none'; // Скрываем контейнер, если прицеп не выбран
            }
        }


        $(document).ready(function() {
            $('#recipient').change(function() {
                const selectedOption = $(this).find(':selected');
                const selectedInn = selectedOption.data('inn');
                const selectedRazgruzka = selectedOption.data('razgruzka');

                $('#hiddenFields').html(`
                    <input type="hidden" name="inn" value="${selectedInn}">
                    <input type="hidden" name="razgruzka" value="${selectedRazgruzka}">
                `);
            });

            loadTrailers();

            $('#senders').change(function() {
                const senderId = $(this).val();
                if (senderId) {
                    $.ajax({
                        url: '/get_addresses/' + senderId,
                        type: 'GET',
                        success: function(data) {
                            $('#addresses').empty();
                            $.each(data, function(i, address) {
                                $('#addresses').append('<option value="' + address[0] + '">' + address[1] + '</option>');
                            });
                        }
                    });
                } else {
                    $('#addresses').empty();
                }
            });

            // Включение select2 с возможностью добавления новых значений
            $('.select2-taggable').select2({
                    tags: true
                });

            $('#delivery_method, #raw_material, #drivers, #transport, #laboratory, #recipient').select2({ tags: true });
        });