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
    #       { operation: 'greaterThan', val: 5 },
    #       { operation: 'lessThanOrEqual', val: 10 }
    #     ]
    #   )
    def initialize(options={})
      parse_options options
    end

    serializable_attributes :and

    def apply(cell)
      # no-op?
      # do we need to manually apply the filters here? that will be lame, if so.
    end

    attr_reader :and

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
      @custom_filter_items ||= []
    end

    def custom_filter_items=(values)
      values.each do |value|
        custom_filter_items << CustomFilter.new(value)
      end
    end

    class CustomFilter
      include Axlsx::OptionsParser

      def initialize(options={})
        raise ArgumentError, "You must specify an operator for the custom filter" unless options[:operator]
        raise ArgumentError, "You must specify a val for the custom filter" unless options[:val]
        parse_options options
      end

      attr_reader :operator
      attr_accessor :val

      def operator=(operation_type)
        RestrictionValidator.validate "CustomFilters.operator", OPERATION_TYPES, operation_type
        @operator = operation_type
      end

      # Serializes the custom_filter object
      # @param [String] str The string to concat the serialization information to.
      def to_xml_string(str = '')
        str << "<customFilter operation='#{@operator}' val='#{@val.to_s}'"
      end
    end
  end
end
