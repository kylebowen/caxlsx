module Axlsx
  class CustomFilters
    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # Creates a new CustomFilters object
    # @param [Hash] options Options used to set this objects attributes and
    #                       create custom_filter_items.
    # @option [Boolean] and @see and.
    # @option [Array] custom_filter_items An array of values that will be used to create custom_filter objects. TODO
    # @note The recommended way to interact with custom_filter objects is via AutoFilter#add_column
    # @example
    #   ws.auto_filter.add_column(
    #     0,
    #     :custom_filters,
    #     and: true,
    #     custom_filter_items: [
    #       { operator: 'greaterThan', val: 5 },
    #       { operator: 'lessThanOrEqual', val: 10 }
    #     ]
    #   )
    def initialize(options={})
      parse_options options
    end

    serializable_attributes :and

    # Tells us if the row of the cell provided should be hidden as it
    # does not meet any or all (based on the logical_and boolean) of
    # the specified custom_filter_items restrictions .
    # @param [Cell] cell The cell to test against items
    def apply(cell)
      return false unless cell

      if logical_and
        # false = show because it matched all criteria
        # true = hide because it failed to match one or both criteria
        return custom_filter_items.any? { |custom_filter| custom_filter.apply(cell) }
      else
        # false = show because it matched at least one criteria
        # true = hide because it didn't match any criteria
        return custom_filter_items.none? { |custom_filter| custom_filter.apply(cell) }
      end

      true
    end

    attr_reader :and
    alias_method :logical_and, :and

    def and=(bool)
      Axlsx.validate_boolean bool
      @and = bool
    end

    def to_xml_string(str = '')
      str << "<customFilters #{serialized_attributes}>"
      custom_filter_items.each { |custom_filter| custom_filter.to_xml_string(str) }
      str << "</customFilters>"
    end

    def custom_filter_items
      @custom_filter_items ||= SimpleTypedList.new(CustomFilter)
    end

    def custom_filter_items=(values)
      values.each do |value|
        custom_filter_items << CustomFilter.new(value)
      end
    end

    class CustomFilter
      include Axlsx::OptionsParser

      OPERATOR_MAP = {
        "equal" => :==,
        "greaterThan" => :>,
        "greaterThanOrEqual" => :>=,
        "lessThan" => :<,
        "lessThanOrEqual" => :<=,
        "notEqual" => :!=,
      }

      def initialize(options={})
        raise ArgumentError, "You must specify an operator for the custom filter" unless options[:operator]
        raise ArgumentError, "You must specify a val for the custom filter" unless options[:val]
        parse_options options
      end

      attr_reader :operator
      attr_accessor :val

      def operator=(operation_type)
        RestrictionValidator.validate "CustomFilter.operator", OPERATOR_MAP.keys, operation_type
        @operator = operation_type
      end

      def apply(cell)
        return false unless cell
        return false if cell.value.send(OPERATOR_MAP[operator], val)

        true
      end

      # Serializes the custom_filter object
      # @param [String] str The string to concat the serialization information to.
      def to_xml_string(str = '')
        str << "<customFilter operator='#{@operator}' val='#{@val.to_s}'"
      end
    end
  end
end
