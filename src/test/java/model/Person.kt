package model

import java.math.BigDecimal
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.*

data class Person(
        var name: String = "",
        var age: Int = 0,
        var weight: Double = 0.0,
        var height: Float = 0f,
        var wage: BigDecimal = BigDecimal.ZERO,
        var birthDayOld: Date? = null,
        var birthDayLocalDate: LocalDate? = null,
        var birthDayLocalDateTime: LocalDateTime? = null,
        var address: Address = Address()
)