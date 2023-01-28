module Axlsx
  class CustomFilters
    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # Creates a new CustomFilters object
    # @param [Hash] options Options used to set this objects attributes and
    #                       create custom_filter_items.
    # @option [Boolean] and @see and.
    # @option [Array] custom_filter_items An array of values that will be used to create custom_filter objects.
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
      @and = false
      parse_options options
    end

    serializable_attributes :and

    attr_reader :and
    alias_method :logical_and, :and

    def and=(bool)
      Axlsx.validate_boolean bool
      @and = bool
    end

    # Tells us if the row of the cell provided should be hidden as it
    # does not meet any or all (based on the logical_and boolean) of
    # the specified custom_filter_items restrictions.
    # @param [Cell] cell The cell to test against items
    def apply(cell)
      return false unless cell

      if logical_and
        custom_filter_items.all? { |custom_filter| custom_filter.apply(cell) }
      else
        custom_filter_items.any? { |custom_filter| custom_filter.apply(cell) }
      end
    end

    # Serialize the object to xml
    # @param [String] str The string to concat the serialization information to.
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
        if value[:operator] == "between"
          custom_filter_items << CustomFilter.new(comparator: "greaterThanOrEqual", operator: "greaterThanOrEqual", val: value[:val][0])
          custom_filter_items << CustomFilter.new(comparator: "lessThanOrEqual", operator: "lessThanOrEqual", val: value[:val][1])
        else
          custom_filter_items << CustomFilter.new(value)
        end
      end
    end

    class CustomFilter
      include Axlsx::OptionsParser

      COMPARATOR_MAP = {
        "lessThan" => :<,
        "lessThanOrEqual" => :<=,
        "equal" => :==,
        "notBlank" => :!=,
        "notEqual" => :!=,
        "greaterThanOrEqual" => :>=,
        "greaterThan" => :>,
        "contains" => :include?,
        "notContains" => :exclude?,
        "beginsWith" => :starts_with?,
        "endsWith" => :ends_with?,
      }
      OPERATORS = ["equal", "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual", "notEqual"]

      def initialize(options={})
        raise ArgumentError, "You must specify a comparator for the custom filter" unless options[:comparator]
        raise ArgumentError, "You must specify an operator for the custom filter" unless options[:operator]
        parse_options options
      end

      attr_reader :comparator
      attr_reader :operator
      attr_accessor :val

      def operator=(operation_type)
        RestrictionValidator.validate "CustomFilter.operator", OPERATORS, operation_type
        @operator = operation_type
      end

      def comparator=(comparator_type)
        RestrictionValidator.validate "CustomFilter.comparator", COMPARATOR_MAP.keys, comparator_type
        @comparator = comparator_type
      end

      def apply(cell)
        return false unless cell
        return false if cell.value.send(COMPARATOR_MAP[comparator], val)

        true
      end

      # Serializes the custom_filter object
      # @param [String] str The string to concat the serialization information to.
      def to_xml_string(str = '')
        str << "<customFilter operator=\"#{@operator}\" val=\"#{leading_wildcard + safe_val + trailing_wildcard}\" />"
      end

      private

      def leading_wildcard
        ["contains", "notContains", "endsWith"].include?(comparator) ? "*" : ""
      end

      def safe_val
        if comparator == "notBlank" && @val.nil?
          " "
        else
          @val.to_s
        end
      end

      def trailing_wildcard
        ["contains", "notContains", "beginsWith"].include?(comparator) ? "*" : ""
      end
    end
  end
end
